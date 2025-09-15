import logging
import os
import time
import re
from typing import Optional, List, Tuple
from openai import OpenAI, AsyncOpenAI

from open_webui.env import SRC_LOG_LEVELS
from open_webui.retrieval.models.base_reranker import BaseReranker

log = logging.getLogger(__name__)
log.setLevel(SRC_LOG_LEVELS["RAG"])


class OpenAIReranker(BaseReranker):
    def __init__(
        self,
        api_key: str,
        base_url: str,
        model: str = "gpt-4o",
        temperature: float = 0.0,
        use_rankgpt: bool = True,
        max_length: int = 300,
    ):
        """
        Initialize the OpenAI Reranker.
        
        Args:
            api_key: OpenAI API key
            model: OpenAI model to use for reranking
            temperature: Temperature for the model
            use_rankgpt: Whether to use RankGPT conversational approach
            max_length: Maximum length of content per passage
        """
        if not api_key:
            raise ValueError("OpenAI API key is required but not provided")
        
        # self.base_url = base_url + "api_version=2024-12-01-preview"

        self.client = OpenAI(api_key=api_key, base_url=base_url)
        
        self.api_key = api_key
        self.model = model
        self.base_url = base_url
        self.temperature = temperature
        self.use_rankgpt = use_rankgpt
        self.max_length = max_length
        
        log.info(f"OpenAI Reranker initialized with model: {self.model}, use_rankgpt: {self.use_rankgpt}")

    def _get_prefix_prompt(self, query: str, num: int) -> List[dict]:
        """Get the prefix prompt for RankGPT approach."""
        return [
            {
                'role': 'system',
                'content': "You are RankGPT, an intelligent assistant that can rank passages based on their relevancy to the query."
            },
            {
                'role': 'user',
                'content': f"I will provide you with {num} passages, each indicated by number identifier []. \nRank the passages based on their relevance to query: {query}."
            },
            {
                'role': 'assistant',
                'content': 'Okay, please provide the passages.'
            }
        ]

    def _get_post_prompt(self, query: str, num: int) -> str:
        """Get the post prompt for RankGPT approach."""
        return f"Search Query: {query}. \nRank the {num} passages above based on their relevance to the search query. The passages should be listed in descending order using identifiers. The most relevant passages should be listed first. The output format should be [] > [], e.g., [1] > [2]. Only response the ranking results, do not say any word or explain."

    def _create_permutation_instruction(self, query: str, documents: List[str]) -> List[dict]:
        """Create the permutation instruction for RankGPT approach."""
        num = len(documents)
        messages = self._get_prefix_prompt(query, num)
        
        for rank, content in enumerate(documents, 1):
            # Clean and truncate content
            content = content.replace('Title: Content: ', '').strip()
            # Truncate by words to respect max_length
            content = ' '.join(content.split()[:self.max_length])
            
            messages.append({
                'role': 'user',
                'content': f"[{rank}] {content}"
            })
            messages.append({
                'role': 'assistant',
                'content': f'Received passage [{rank}].'
            })
        
        messages.append({
            'role': 'user',
            'content': self._get_post_prompt(query, num)
        })
        
        return messages

    def _parse_rankgpt_response(self, response: str, num_docs: int) -> List[int]:
        """Parse RankGPT response to extract ranking order."""
        try:
            # Extract ranking pattern like [1] > [2] > [3]
            pattern = r'\[(\d+)\]'
            matches = re.findall(pattern, response)
            
            if not matches:
                log.warning("No ranking pattern found in response, using original order")
                return list(range(1, num_docs + 1))
            
            # Convert to 0-based indices
            ranking = [int(match) - 1 for match in matches]
            
            # Validate ranking
            if len(ranking) != num_docs or set(ranking) != set(range(num_docs)):
                log.warning("Invalid ranking received, using original order")
                return list(range(num_docs))
            
            return ranking
            
        except Exception as e:
            log.error(f"Error parsing RankGPT response: {e}")
            return list(range(num_docs))

    def _rerank_with_rankgpt(self, query: str, documents: List[str]) -> List[Tuple[str, float]]:
        """Rerank documents using RankGPT approach."""

        try:
            messages = self._create_permutation_instruction(query, documents)
            
            response = self.client.chat.completions.create(
                model=self.model,
                messages=messages,
                temperature=self.temperature,
                # max_tokens=1000
            )
            
            response_text = response.choices[0].message.content.strip()
            log.debug(f"RankGPT response: {response_text}")
            
            # Parse the ranking
            ranking = self._parse_rankgpt_response(response_text, len(documents))
            
            # Reorder documents based on ranking
            reranked_docs = []
            for i, doc_idx in enumerate(ranking):
                # Calculate score based on position (higher position = higher score)
                score = (len(documents) - i) / len(documents) * 1.000000
                reranked_docs.append((documents[doc_idx], score))
            
            return reranked_docs
            
        except Exception as e:
            log.error(f"Error in RankGPT reranking: {e}")
            # Fallback to original order with equal scores
            return [(doc, 0.5) for doc in documents]

    def _rerank_with_json(self, query: str, documents: List[str]) -> List[Tuple[str, float]]:
        """Rerank documents using JSON format approach."""
        try:
            # Create documents list for the prompt
            docs_text = "\n".join([f"Document {i+1}: {doc}" for i, doc in enumerate(documents)])
            
            messages = [
                {
                    "role": "system",
                    "content": "You are an expert relevance ranker. Given a list of documents and a query, your job is to determine how relevant each document is for answering the query. Your output is JSON, which is a list of documents. Each document has two fields, content and score. relevance_score is from 0.000000 to 1.000000. Higher relevance means higher score."
                },
                {
                    "role": "user",
                    "content": f"Query: {query}\nDocs: {docs_text}"
                }
            ]
            
            response = self.client.chat.completions.create(
                model=self.model,
                response_format={"type": "json_object"},
                messages=messages,
                temperature=self.temperature
            )
            
            response_text = response.choices[0].message.content.strip()
            result = json.loads(response_text)
            
            # Extract documents and scores
            reranked_docs = []
            for doc_data in result.get("documents", []):
                content = doc_data.get("content", "")
                score = doc_data.get("score", 0.0)
                reranked_docs.append((content, score))
            
            return reranked_docs
            
        except Exception as e:
            log.error(f"Error in JSON reranking: {e}")
            # Fallback to original order with equal scores
            return [(doc, 0.5) for doc in documents]

    def rerank(self, query: str, documents: List[str]) -> List[Tuple[str, float]]:
        """
        Rerank documents based on relevance to the query.
        
        Args:
            query: The search query
            documents: List of document contents to rerank
            
        Returns:
            List of tuples (document_content, relevance_score)
        """
        if not documents:
            return []
        
        start_time = time.time()
        
        try:
            if self.use_rankgpt:
                reranked_docs = self._rerank_with_rankgpt(query, documents)
            else:
                reranked_docs = self._rerank_with_json(query, documents)
            
            elapsed_time = time.time() - start_time
            log.info(f"Reranked {len(documents)} documents in {elapsed_time:.2f} seconds using {self.model}")
            
            return reranked_docs
            
        except Exception as e:
            log.error(f"Error in OpenAI reranking: {e}")
            # Fallback to original order with equal scores
            return [(doc, 0.5) for doc in documents]

    def predict(self, sentences: List[Tuple[str, str]]) -> Optional[List[float]]:
        """
        Predict relevance scores for sentence pairs.
        This method is required by the BaseReranker interface.
        
        Args:
            sentences: List of (query, document) tuples
            
        Returns:
            List of relevance scores
        """
        if not sentences:
            return []
        
        try:
            # Extract documents from the sentence pairs
            documents = [doc for query, doc in sentences]
            query = sentences[0][0] if sentences else ""

            reranked_docs = self.rerank(query, documents)
            
            # Create a fast lookup dictionary for scores
            doc_to_score = {doc: score for doc, score in reranked_docs}
            
            # Extract scores in the same order as input sentences (no conversion needed)
            scores = [doc_to_score.get(doc, 0.5) for query, doc in sentences]
            
            return scores
            
        except Exception as e:
            log.error(f"Error in OpenAI reranker predict: {e}")
            # Return equal scores as fallback
            return [0.5] * len(sentences)

    def get_model_name(self) -> str:
        """Get the model name."""
        return f"openai-{self.model}"