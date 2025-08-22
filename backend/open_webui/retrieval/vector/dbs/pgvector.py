import io
import time
from typing import Optional, List, Dict, Any
import logging
import json
from sqlalchemy import (
    func,
    literal,
    cast,
    column,
    create_engine,
    Column,
    Integer,
    MetaData,
    LargeBinary,
    select,
    text,
    Text,
    Table,
    values,
)
from sqlalchemy.sql import true
from sqlalchemy.pool import NullPool, QueuePool

from sqlalchemy.orm import declarative_base, scoped_session, sessionmaker
from sqlalchemy.dialects.postgresql import JSONB, array
from pgvector.sqlalchemy import Vector
from sqlalchemy.ext.mutable import MutableDict
from sqlalchemy.exc import NoSuchTableError


from open_webui.retrieval.vector.utils import stringify_metadata
from open_webui.retrieval.vector.main import (
    VectorDBBase,
    VectorItem,
    SearchResult,
    GetResult,
)
from open_webui.config import (
    PGVECTOR_DB_URL,
    PGVECTOR_INITIALIZE_MAX_VECTOR_LENGTH,
    PGVECTOR_PGCRYPTO,
    PGVECTOR_PGCRYPTO_KEY,
    PGVECTOR_POOL_SIZE,
    PGVECTOR_POOL_MAX_OVERFLOW,
    PGVECTOR_POOL_TIMEOUT,
    PGVECTOR_POOL_RECYCLE,
)

from open_webui.env import SRC_LOG_LEVELS

VECTOR_LENGTH = PGVECTOR_INITIALIZE_MAX_VECTOR_LENGTH
Base = declarative_base()

log = logging.getLogger(__name__)
log.setLevel(SRC_LOG_LEVELS["RAG"])


def pgcrypto_encrypt(val, key):
    return func.pgp_sym_encrypt(val, literal(key))


def pgcrypto_decrypt(col, key, outtype="text"):
    return func.cast(func.pgp_sym_decrypt(col, literal(key)), outtype)

class VectorCopyInsert:
    def __init__(self, session):
        self.session = session
    
    def copy_bulk_insert(self, items, collection_name):
        """
        Ultra-fast bulk insert using PostgreSQL COPY for large vector tables
        This bypasses SQLAlchemy and uses raw COPY for maximum performance
        """
        start_time = time.time()
        log.info(f"Starting COPY bulk insert of {len(items)} items")
        
        # Step 1: Prepare data in memory (much faster than file I/O)
        copy_buffer = io.StringIO()
        
        for item in items:
            # Batch process vector adjustments
            vector = self.adjust_vector_length(item["vector"])
            
            # Convert vector to PostgreSQL array format: {1.0,2.0,3.0}
            vector_str = '{' + ','.join(f'{v:.6f}' for v in vector) + '}'
            
            # Handle metadata properly
            if isinstance(item["metadata"], dict):
                metadata_json = json.dumps(item["metadata"])
            else:
                metadata_json = str(item["metadata"])
            
            # Escape text for COPY (handle newlines, tabs, backslashes)
            text_clean = item["text"].replace('\\', '\\\\').replace('\n', '\\n').replace('\r', '\\r').replace('\t', '\\t')
            metadata_clean = metadata_json.replace('\\', '\\\\').replace('\n', '\\n').replace('\r', '\\r').replace('\t', '\\t')
            
            # Write tab-separated values
            copy_buffer.write(f"{item['id']}\t{vector_str}\t{collection_name}\t{text_clean}\t{metadata_clean}\n")
        
        prep_time = time.time() - start_time
        log.info(f"Data preparation completed in {prep_time:.2f} seconds")
        
        # Step 2: Execute COPY command
        copy_buffer.seek(0)
        copy_start = time.time()
        
        # Get raw connection for COPY
        raw_connection = self.session.connection().connection
        
        try:
            with raw_connection.cursor() as cursor:
                # Temporarily disable autocommit for better performance
                old_autocommit = raw_connection.autocommit
                raw_connection.autocommit = False
                
                log.info("Executing COPY command...")
                cursor.copy_from(
                    copy_buffer,
                    'document_chunk',
                    columns=('id', 'vector', 'collection_name', 'text', 'vmetadata'),
                    sep='\t',
                    null='\\N'  # Handle NULL values properly
                )
                
                # Commit the transaction
                raw_connection.commit()
                raw_connection.autocommit = old_autocommit
                
                copy_time = time.time() - copy_start
                total_time = time.time() - start_time
                
                log.info(f"COPY insert completed in {copy_time:.2f} seconds")
                log.info(f"Total operation time: {total_time:.2f} seconds (prep: {prep_time:.2f}s, copy: {copy_time:.2f}s)")
                log.info(f"Successfully inserted {len(items)} items using COPY")
                
                return True
                
        except Exception as e:
            raw_connection.rollback()
            log.error(f"COPY operation failed: {e}")
            log.info("Falling back to regular INSERT method...")
            return False
        
        finally:
            copy_buffer.close()
    
    def copy_with_conflict_handling(self, items, collection_name):
        """
        COPY with conflict handling using temp table approach
        Since COPY doesn't support ON CONFLICT directly
        """
        import uuid
        temp_table = f"temp_bulk_insert_{uuid.uuid4().hex[:8]}"
        
        try:
            # Create temporary table
            self.session.execute(text(f"""
                CREATE TEMPORARY TABLE {temp_table} (
                    LIKE document_chunk INCLUDING DEFAULTS
                );
            """))
            
            # Use COPY to insert into temp table
            if self.copy_bulk_insert_to_table(items, collection_name, temp_table):
                # Move data with conflict resolution
                self.session.execute(text(f"""
                    INSERT INTO document_chunk 
                    SELECT * FROM {temp_table}
                    ON CONFLICT (id) DO NOTHING;
                """))
                self.session.commit()
                log.info(f"Successfully inserted items with conflict handling")
                return True
            else:
                return False
                
        except Exception as e:
            self.session.rollback()
            log.error(f"COPY with conflict handling failed: {e}")
            return False
        finally:
            # Cleanup temp table
            try:
                self.session.execute(text(f"DROP TABLE IF EXISTS {temp_table};"))
            except:
                pass
    
    def copy_bulk_insert_to_table(self, items, collection_name, table_name):
        """Helper method to COPY into a specific table"""
        copy_buffer = io.StringIO()
        
        for item in items:
            vector = self.adjust_vector_length(item["vector"])
            vector_str = '{' + ','.join(f'{v:.6f}' for v in vector) + '}'
            
            metadata_json = json.dumps(item["metadata"]) if isinstance(item["metadata"], dict) else str(item["metadata"])
            
            text_clean = item["text"].replace('\\', '\\\\').replace('\n', '\\n').replace('\r', '\\r').replace('\t', '\\t')
            metadata_clean = metadata_json.replace('\\', '\\\\').replace('\n', '\\n').replace('\r', '\\r').replace('\t', '\\t')
            
            copy_buffer.write(f"{item['id']}\t{vector_str}\t{collection_name}\t{text_clean}\t{metadata_clean}\n")
        
        copy_buffer.seek(0)
        
        raw_connection = self.session.connection().connection
        try:
            with raw_connection.cursor() as cursor:
                cursor.copy_from(
                    copy_buffer,
                    table_name,
                    columns=('id', 'vector', 'collection_name', 'text', 'vmetadata'),
                    sep='\t',
                    null='\\N'
                )
                return True
        except Exception as e:
            log.error(f"COPY to {table_name} failed: {e}")
            return False
        finally:
            copy_buffer.close()



class DocumentChunk(Base):
    __tablename__ = "document_chunk"

    id = Column(Text, primary_key=True)
    vector = Column(Vector(dim=VECTOR_LENGTH), nullable=True)
    collection_name = Column(Text, nullable=False)

    if PGVECTOR_PGCRYPTO:
        text = Column(LargeBinary, nullable=True)
        vmetadata = Column(LargeBinary, nullable=True)
    else:
        text = Column(Text, nullable=True)
        vmetadata = Column(MutableDict.as_mutable(JSONB), nullable=True)


class PgvectorClient(VectorDBBase):
    def __init__(self) -> None:
        log.info(f"Initializing PgvectorClient with PGVECTOR_DB_URL: {PGVECTOR_DB_URL}")
        # if no pgvector uri, use the existing database connection
        if not PGVECTOR_DB_URL:
            from open_webui.internal.db import Session

            self.session = Session
        else:
            if isinstance(PGVECTOR_POOL_SIZE, int):
                if PGVECTOR_POOL_SIZE > 0:
                    engine = create_engine(
                        PGVECTOR_DB_URL,
                        pool_size=PGVECTOR_POOL_SIZE,
                        max_overflow=PGVECTOR_POOL_MAX_OVERFLOW,
                        pool_timeout=PGVECTOR_POOL_TIMEOUT,
                        pool_recycle=PGVECTOR_POOL_RECYCLE,
                        pool_pre_ping=True,
                        poolclass=QueuePool,
                    )
                else:
                    log.info(f"Creating engine with NullPool")
                    engine = create_engine(
                        PGVECTOR_DB_URL, pool_pre_ping=True, poolclass=NullPool
                    )
            else:
                log.info(f"Creating engine with QueuePool")
                engine = create_engine(PGVECTOR_DB_URL, pool_pre_ping=True)

            SessionLocal = sessionmaker(
                autocommit=False, autoflush=False, bind=engine, expire_on_commit=False
            )
            self.session = scoped_session(SessionLocal)

        try:
            # Ensure the pgvector extension is available
            self.session.execute(text("CREATE EXTENSION IF NOT EXISTS vector;"))

            if PGVECTOR_PGCRYPTO:
                # Ensure the pgcrypto extension is available for encryption
                self.session.execute(text("CREATE EXTENSION IF NOT EXISTS pgcrypto;"))

                if not PGVECTOR_PGCRYPTO_KEY:
                    raise ValueError(
                        "PGVECTOR_PGCRYPTO_KEY must be set when PGVECTOR_PGCRYPTO is enabled."
                    )

            # Check vector length consistency
            self.check_vector_length()

            # Create the tables if they do not exist
            # Base.metadata.create_all requires a bind (engine or connection)
            # Get the connection from the session
            connection = self.session.connection()
            Base.metadata.create_all(bind=connection)

            # Create an index on the vector column if it doesn't exist
            self.session.execute(
                text(
                    "CREATE INDEX IF NOT EXISTS idx_document_chunk_vector "
                    "ON document_chunk USING ivfflat (vector vector_cosine_ops) WITH (lists = 100);"
                )
            )
            log.info(f"Creating index on vector column")
            self.session.execute(
                text(
                    "CREATE INDEX IF NOT EXISTS idx_document_chunk_collection_name "
                    "ON document_chunk (collection_name);"
                )
            )
            log.info(f"Creating index on collection name column")
            self.session.commit()
            log.info("Initialization complete.")
        except Exception as e:
            self.session.rollback()
            log.exception(f"Error during initialization: {e}")
            raise

    def check_vector_length(self) -> None:
        """
        Check if the VECTOR_LENGTH matches the existing vector column dimension in the database.
        Raises an exception if there is a mismatch.
        """
        log.info(f"Checking vector length")
        metadata = MetaData()
        try:
            # Attempt to reflect the 'document_chunk' table
            document_chunk_table = Table(
                "document_chunk", metadata, autoload_with=self.session.bind
            )
        except NoSuchTableError:
            # Table does not exist; no action needed
            return

        # Proceed to check the vector column
        log.info(f"Checking vector length")
        if "vector" in document_chunk_table.columns:
            log.info(f"vector column exists")
            vector_column = document_chunk_table.columns["vector"]
            vector_type = vector_column.type
            if isinstance(vector_type, Vector):
                db_vector_length = vector_type.dim
                if db_vector_length != VECTOR_LENGTH:
                    raise Exception(
                        f"VECTOR_LENGTH {VECTOR_LENGTH} does not match existing vector column dimension {db_vector_length}. "
                        "Cannot change vector size after initialization without migrating the data."
                    )
            else:
                raise Exception(
                    "The 'vector' column exists but is not of type 'Vector'."
                )
        else:
            raise Exception(
                "The 'vector' column does not exist in the 'document_chunk' table."
            )

    def adjust_vector_length(self, vector: List[float]) -> List[float]:
        log.info(f"Adjusting vector length from {len(vector)} to {VECTOR_LENGTH}")
        # Adjust vector to have length VECTOR_LENGTH
        current_length = len(vector)
        if current_length < VECTOR_LENGTH:
            # Pad the vector with zeros
            vector += [0.0] * (VECTOR_LENGTH - current_length)
        elif current_length > VECTOR_LENGTH:
            # Truncate the vector to VECTOR_LENGTH
            vector = vector[:VECTOR_LENGTH]
        return vector


    # Process all vectors at once:
    def batch_adjust_vectors(self, items):
        """Adjust all vectors in one operation"""
        vectors = [item["vector"] for item in items]
        # Assuming adjust_vector_length does padding/truncation
        adjusted_vectors = []
        for vector in vectors:
            if len(vector) == VECTOR_LENGTH:
                adjusted_vectors.append(vector)
            elif len(vector) < VECTOR_LENGTH:
                # Pad with zeros
                adjusted_vectors.append(vector + [0.0] * (VECTOR_LENGTH - len(vector)))
            else:
                # Truncate
                adjusted_vectors.append(vector[:VECTOR_LENGTH])
        return adjusted_vectors
    
    def insert(self, collection_name: str, items: List[VectorItem]) -> None:
        try:
            if PGVECTOR_PGCRYPTO:
                for item in items:
                    vector = self.adjust_vector_length(item["vector"])
                    # Use raw SQL for BYTEA/pgcrypto
                    # Ensure metadata is converted to its JSON text representation
                    json_metadata = json.dumps(item["metadata"])
                    self.session.execute(
                        text(
                            """
                            INSERT INTO document_chunk
                            (id, vector, collection_name, text, vmetadata)
                            VALUES (
                                :id, :vector, :collection_name,
                                pgp_sym_encrypt(:text, :key),
                                pgp_sym_encrypt(:metadata_text, :key)
                            )
                            ON CONFLICT (id) DO NOTHING
                        """
                        ),
                        {
                            "id": item["id"],
                            "vector": vector,
                            "collection_name": collection_name,
                            "text": item["text"],
                            "metadata_text": json_metadata,
                            "key": PGVECTOR_PGCRYPTO_KEY,
                        },
                    )
                self.session.commit()
                log.info(f"Encrypted & inserted {len(items)} into '{collection_name}'")

            else:
                log.info(
                    f"Called insert {len(items)} items into collection '{collection_name}'."
                )

                start_time = time.time()
                log.info("Dropping vector index for bulk insert")
                self.session.execute(text("DROP INDEX IF EXISTS idx_document_chunk_vector;"))

                # Use COPY method (handles everything internally)
                log.info(f"Using COPY method for {len(items)} items")
                insert_start = time.time()

                copy_inserter = VectorCopyInsert(self.session)
                success = copy_inserter.copy_bulk_insert(items, collection_name)

                insert_time = time.time() - insert_start
                log.info(f"COPY insert completed in {insert_time:.2f} seconds")

                # Note: No need for session.commit() - copy_bulk_insert handles it

                # Recreate vector index
                log.info("Recreating vector index")
                index_start = time.time()
                self.session.execute(text("""
                    CREATE INDEX idx_document_chunk_vector 
                    ON document_chunk USING ivfflat (vector vector_cosine_ops) 
                    WITH (lists = 200);
                """))

                index_time = time.time() - index_start
                total_time = time.time() - start_time

                log.info(f"Index recreation took {index_time:.2f} seconds")
                log.info(f"Total time: {total_time:.2f}s (insert: {insert_time:.2f}s, index: {index_time:.2f}s)")
                
        except Exception as e:
            self.session.rollback()
            log.exception(f"Error during insert: {e}")
            raise

    def upsert(self, collection_name: str, items: List[VectorItem]) -> None:
        try:
            if PGVECTOR_PGCRYPTO:
                for item in items:
                    vector = self.adjust_vector_length(item["vector"])
                    json_metadata = json.dumps(item["metadata"])
                    self.session.execute(
                        text(
                            """
                            INSERT INTO document_chunk
                            (id, vector, collection_name, text, vmetadata)
                            VALUES (
                                :id, :vector, :collection_name,
                                pgp_sym_encrypt(:text, :key),
                                pgp_sym_encrypt(:metadata_text, :key)
                            )
                            ON CONFLICT (id) DO UPDATE SET
                              vector = EXCLUDED.vector,
                              collection_name = EXCLUDED.collection_name,
                              text = EXCLUDED.text,
                              vmetadata = EXCLUDED.vmetadata
                        """
                        ),
                        {
                            "id": item["id"],
                            "vector": vector,
                            "collection_name": collection_name,
                            "text": item["text"],
                            "metadata_text": json_metadata,
                            "key": PGVECTOR_PGCRYPTO_KEY,
                        },
                    )
                self.session.commit()
                log.info(f"Encrypted & upserted {len(items)} into '{collection_name}'")
            else:
                for item in items:
                    vector = self.adjust_vector_length(item["vector"])
                    existing = (
                        self.session.query(DocumentChunk)
                        .filter(DocumentChunk.id == item["id"])
                        .first()
                    )
                    if existing:
                        existing.vector = vector
                        existing.text = item["text"]
                        existing.vmetadata = stringify_metadata(item["metadata"])
                        existing.collection_name = (
                            collection_name  # Update collection_name if necessary
                        )
                    else:
                        new_chunk = DocumentChunk(
                            id=item["id"],
                            vector=vector,
                            collection_name=collection_name,
                            text=item["text"],
                            vmetadata=stringify_metadata(item["metadata"]),
                        )
                        self.session.add(new_chunk)
                self.session.commit()
                log.info(
                    f"Upserted {len(items)} items into collection '{collection_name}'."
                )
        except Exception as e:
            self.session.rollback()
            log.exception(f"Error during upsert: {e}")
            raise

    def search(
        self,
        collection_name: str,
        vectors: List[List[float]],
        limit: Optional[int] = None,
    ) -> Optional[SearchResult]:
        try:
            if not vectors:
                return None

            # Adjust query vectors to VECTOR_LENGTH
            vectors = [self.adjust_vector_length(vector) for vector in vectors]
            num_queries = len(vectors)

            def vector_expr(vector):
                return cast(array(vector), Vector(VECTOR_LENGTH))

            # Create the values for query vectors
            qid_col = column("qid", Integer)
            q_vector_col = column("q_vector", Vector(VECTOR_LENGTH))
            query_vectors = (
                values(qid_col, q_vector_col)
                .data(
                    [(idx, vector_expr(vector)) for idx, vector in enumerate(vectors)]
                )
                .alias("query_vectors")
            )

            result_fields = [
                DocumentChunk.id,
            ]
            if PGVECTOR_PGCRYPTO:
                result_fields.append(
                    pgcrypto_decrypt(
                        DocumentChunk.text, PGVECTOR_PGCRYPTO_KEY, Text
                    ).label("text")
                )
                result_fields.append(
                    pgcrypto_decrypt(
                        DocumentChunk.vmetadata, PGVECTOR_PGCRYPTO_KEY, JSONB
                    ).label("vmetadata")
                )
            else:
                result_fields.append(DocumentChunk.text)
                result_fields.append(DocumentChunk.vmetadata)
            result_fields.append(
                (DocumentChunk.vector.cosine_distance(query_vectors.c.q_vector)).label(
                    "distance"
                )
            )

            # Build the lateral subquery for each query vector
            subq = (
                select(*result_fields)
                .where(DocumentChunk.collection_name == collection_name)
                .order_by(
                    (DocumentChunk.vector.cosine_distance(query_vectors.c.q_vector))
                )
            )
            if limit is not None:
                subq = subq.limit(limit)
            subq = subq.lateral("result")

            # Build the main query by joining query_vectors and the lateral subquery
            stmt = (
                select(
                    query_vectors.c.qid,
                    subq.c.id,
                    subq.c.text,
                    subq.c.vmetadata,
                    subq.c.distance,
                )
                .select_from(query_vectors)
                .join(subq, true())
                .order_by(query_vectors.c.qid, subq.c.distance)
            )

            result_proxy = self.session.execute(stmt)
            results = result_proxy.all()

            ids = [[] for _ in range(num_queries)]
            distances = [[] for _ in range(num_queries)]
            documents = [[] for _ in range(num_queries)]
            metadatas = [[] for _ in range(num_queries)]

            if not results:
                return SearchResult(
                    ids=ids,
                    distances=distances,
                    documents=documents,
                    metadatas=metadatas,
                )

            for row in results:
                qid = int(row.qid)
                ids[qid].append(row.id)
                # normalize and re-orders pgvec distance from [2, 0] to [0, 1] score range
                # https://github.com/pgvector/pgvector?tab=readme-ov-file#querying
                distances[qid].append((2.0 - row.distance) / 2.0)
                documents[qid].append(row.text)
                metadatas[qid].append(row.vmetadata)

            self.session.rollback()  # read-only transaction
            return SearchResult(
                ids=ids, distances=distances, documents=documents, metadatas=metadatas
            )
        except Exception as e:
            self.session.rollback()
            log.exception(f"Error during search: {e}")
            return None

    def query(
        self, collection_name: str, filter: Dict[str, Any], limit: Optional[int] = None
    ) -> Optional[GetResult]:
        try:
            if PGVECTOR_PGCRYPTO:
                # Build where clause for vmetadata filter
                where_clauses = [DocumentChunk.collection_name == collection_name]
                for key, value in filter.items():
                    # decrypt then check key: JSON filter after decryption
                    where_clauses.append(
                        pgcrypto_decrypt(
                            DocumentChunk.vmetadata, PGVECTOR_PGCRYPTO_KEY, JSONB
                        )[key].astext
                        == str(value)
                    )
                stmt = select(
                    DocumentChunk.id,
                    pgcrypto_decrypt(
                        DocumentChunk.text, PGVECTOR_PGCRYPTO_KEY, Text
                    ).label("text"),
                    pgcrypto_decrypt(
                        DocumentChunk.vmetadata, PGVECTOR_PGCRYPTO_KEY, JSONB
                    ).label("vmetadata"),
                ).where(*where_clauses)
                if limit is not None:
                    stmt = stmt.limit(limit)
                results = self.session.execute(stmt).all()
            else:
                query = self.session.query(DocumentChunk).filter(
                    DocumentChunk.collection_name == collection_name
                )

                for key, value in filter.items():
                    query = query.filter(
                        DocumentChunk.vmetadata[key].astext == str(value)
                    )

                if limit is not None:
                    query = query.limit(limit)

                results = query.all()

            if not results:
                return None

            ids = [[result.id for result in results]]
            documents = [[result.text for result in results]]
            metadatas = [[result.vmetadata for result in results]]

            self.session.rollback()  # read-only transaction
            return GetResult(
                ids=ids,
                documents=documents,
                metadatas=metadatas,
            )
        except Exception as e:
            self.session.rollback()
            log.exception(f"Error during query: {e}")
            return None

    def get(
        self, collection_name: str, limit: Optional[int] = None
    ) -> Optional[GetResult]:
        try:
            if PGVECTOR_PGCRYPTO:
                stmt = select(
                    DocumentChunk.id,
                    pgcrypto_decrypt(
                        DocumentChunk.text, PGVECTOR_PGCRYPTO_KEY, Text
                    ).label("text"),
                    pgcrypto_decrypt(
                        DocumentChunk.vmetadata, PGVECTOR_PGCRYPTO_KEY, JSONB
                    ).label("vmetadata"),
                ).where(DocumentChunk.collection_name == collection_name)
                if limit is not None:
                    stmt = stmt.limit(limit)
                results = self.session.execute(stmt).all()
                ids = [[row.id for row in results]]
                documents = [[row.text for row in results]]
                metadatas = [[row.vmetadata for row in results]]
            else:

                query = self.session.query(DocumentChunk).filter(
                    DocumentChunk.collection_name == collection_name
                )
                if limit is not None:
                    query = query.limit(limit)

                results = query.all()

                if not results:
                    return None

                ids = [[result.id for result in results]]
                documents = [[result.text for result in results]]
                metadatas = [[result.vmetadata for result in results]]

            self.session.rollback()  # read-only transaction
            return GetResult(ids=ids, documents=documents, metadatas=metadatas)
        except Exception as e:
            self.session.rollback()
            log.exception(f"Error during get: {e}")
            return None

    def delete(
        self,
        collection_name: str,
        ids: Optional[List[str]] = None,
        filter: Optional[Dict[str, Any]] = None,
    ) -> None:
        try:
            if PGVECTOR_PGCRYPTO:
                wheres = [DocumentChunk.collection_name == collection_name]
                if ids:
                    wheres.append(DocumentChunk.id.in_(ids))
                if filter:
                    for key, value in filter.items():
                        wheres.append(
                            pgcrypto_decrypt(
                                DocumentChunk.vmetadata, PGVECTOR_PGCRYPTO_KEY, JSONB
                            )[key].astext
                            == str(value)
                        )
                stmt = DocumentChunk.__table__.delete().where(*wheres)
                result = self.session.execute(stmt)
                deleted = result.rowcount
            else:
                query = self.session.query(DocumentChunk).filter(
                    DocumentChunk.collection_name == collection_name
                )
                if ids:
                    query = query.filter(DocumentChunk.id.in_(ids))
                if filter:
                    for key, value in filter.items():
                        query = query.filter(
                            DocumentChunk.vmetadata[key].astext == str(value)
                        )
                deleted = query.delete(synchronize_session=False)
            self.session.commit()
            log.info(f"Deleted {deleted} items from collection '{collection_name}'.")
        except Exception as e:
            self.session.rollback()
            log.exception(f"Error during delete: {e}")
            raise

    def reset(self) -> None:
        try:
            deleted = self.session.query(DocumentChunk).delete()
            self.session.commit()
            log.info(
                f"Reset complete. Deleted {deleted} items from 'document_chunk' table."
            )
        except Exception as e:
            self.session.rollback()
            log.exception(f"Error during reset: {e}")
            raise

    def close(self) -> None:
        pass

    def has_collection(self, collection_name: str) -> bool:
        try:
            exists = (
                self.session.query(DocumentChunk)
                .filter(DocumentChunk.collection_name == collection_name)
                .first()
                is not None
            )
            self.session.rollback()  # read-only transaction
            return exists
        except Exception as e:
            self.session.rollback()
            log.exception(f"Error checking collection existence: {e}")
            return False

    def delete_collection(self, collection_name: str) -> None:
        self.delete(collection_name)
        log.info(f"Collection '{collection_name}' deleted.")
