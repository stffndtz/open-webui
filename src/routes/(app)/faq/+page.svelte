<script>
	import { getContext, onMount } from 'svelte';
	import { goto } from '$app/navigation';

	const i18n = getContext('i18n');

	import { mobile, showSidebar, user, WEBUI_NAME } from '$lib/stores';

	import Tooltip from '$lib/components/common/Tooltip.svelte';
	import Sidebar from '$lib/components/icons/Sidebar.svelte';
	import Search from '$lib/components/icons/Search.svelte';
	import XMark from '$lib/components/icons/XMark.svelte';
	import ChevronDown from '$lib/components/icons/ChevronDown.svelte';

	let loaded = false;
	let query = '';
	let selectedCategory = '';

	// FAQ data from Laila Knowledge base
	const faqItems = [
		{
			id: 1,
			category: 'getting-started',
			question: 'How do I start a new chat?',
			answer: `Click "New Chat" (or the "+" button) to begin a fresh conversation with Laila. You can choose which model to use before sending your first message. Each chat is independent unless you explicitly share or copy content between them. Each chat is private unless you explicitly share it with others.`
		},
		{
			id: 2,
			category: 'features',
			question: 'How do I use web search?',
			answer: `You can prefix a URL with # (or start your prompt with #) to fetch and incorporate external web content into the chat. In a chat you can activate the web search button so the AI can search online for up-to-date information.`
		},
		{
			id: 3,
			category: 'features',
			question: 'How do I use a saved prompt?',
			answer: `Use the / command in the chat input to open prompt presets (saved prompts).`
		},
		{
			id: 4,
			category: 'features',
			question: 'How do I save and share a prompt?',
			answer: `In your workspace (German: Arbeitsbereich), go to "Prompts". In the "Prompts" area, press the "+" and define a new prompt template. Under "access" you can share the prompt with your team or the entire organization.`
		},
		{
			id: 5,
			category: 'knowledge',
			question: 'How do I create and share knowledge?',
			answer: `In your workspace (German: "Arbeitsbereich"), go to "Knowledge". There you can find and add Knowledges. Under the Knowledge tab, you can create a new knowledge by pressing "+" and upload files (PDFs, Word, Excel, PPT, etc.) for Laila to "know". Pressing "Access" allows you to share knowledge with your team (choose "Private" and then a group) or the entire organization ("Public").`
		},
		{
			id: 6,
			category: 'knowledge',
			question: 'How do I use knowledge in a chat?',
			answer: `Within a chat, reference knowledge via the # command or by pointing to a document (e.g., #MyDoc.pdf) to bring its content into the context. The system will retrieve relevant passages from the knowledge base to inform the AI's answers. This helps give more accurate, context-aware responses.`
		},
		{
			id: 7,
			category: 'organization',
			question: 'How do I organize my chats in folders?',
			answer: `You can create folders and drag & drop chats into them for organization (press "+" left of the chat label). This helps keep your workspace tidy, especially as you have many conversations on different topics.`
		},
		{
			id: 8,
			category: 'sharing',
			question: 'How do I share a chat?',
			answer: `You can generate shareable chat links to let others view or continue the conversation. Others will get a "snapshot" of the chat and will not be able to see any updates from the chat after sharing. Others can "clone" a chat and continue working on the chat results in their own account.`
		},
		{
			id: 9,
			category: 'organization',
			question: 'How does chat history, archiving and exporting work?',
			answer: `New chats are saved in history by default (unless chat history is disabled). You can archive chats to remove them from the main view. Chats can be exported as JSON, PDF, or TXT for backups or sharing. Importing chats is supported - just drag a JSON file into the sidebar.`
		},
		{
			id: 10,
			category: 'features',
			question: 'How does memory work across chats?',
			answer: `Under Settings > Personalization, you can manually add "Memories" (facts or information you want the AI to remember between chats).`
		},
		{
			id: 11,
			category: 'features',
			question: 'How do I switch modes or use multiple models?',
			answer: `You can switch the model mid-chat via the "+" in the mode dropdown (top left). The system supports multiple models in parallel (many-model chats), and can merge responses. You can even have multiple instances of the same model in a chat.`
		},
		{
			id: 12,
			category: 'settings',
			question: 'How do I customize the interface?',
			answer: `In your personal settings under "General", you can select between different themes: light, dark, OLED dark modes, and custom chat backgrounds.`
		}
	];

	const categories = [
		{ id: 'getting-started', label: 'Getting Started' },
		{ id: 'features', label: 'Features' },
		{ id: 'knowledge', label: 'Knowledge' },
		{ id: 'sharing', label: 'Sharing' },
		{ id: 'organization', label: 'Organization' },
		{ id: 'settings', label: 'Settings' }
	];

	$: filteredItems = faqItems.filter((item) => {
		const matchesQuery =
			query === '' ||
			item.question.toLowerCase().includes(query.toLowerCase()) ||
			item.answer.toLowerCase().includes(query.toLowerCase());
		const matchesCategory = selectedCategory === '' || item.category === selectedCategory;
		return matchesQuery && matchesCategory;
	});

	onMount(async () => {
		// Admin-only access
		if ($user?.role !== 'admin') {
			await goto('/');
			return;
		}
		loaded = true;
	});
</script>

<svelte:head>
	<title>{$i18n.t('FAQ')} | {$WEBUI_NAME}</title>
</svelte:head>

{#if loaded}
	<div
		class="flex flex-col w-full h-screen max-h-[100dvh] transition-width duration-200 ease-in-out {$showSidebar
			? 'md:max-w-[calc(100%-260px)]'
			: ''} max-w-full"
	>
		<!-- Header -->
		<nav class="px-2 pt-1.5 backdrop-blur-xl w-full drag-region">
			<div class="flex items-center">
				{#if $mobile}
					<div class="{$showSidebar ? 'md:hidden' : ''} flex flex-none items-center">
						<Tooltip
							content={$showSidebar ? $i18n.t('Close Sidebar') : $i18n.t('Open Sidebar')}
						>
							<button
								class="cursor-pointer flex rounded-lg hover:bg-gray-100 dark:hover:bg-gray-850 transition"
								on:click={() => {
									showSidebar.set(!$showSidebar);
								}}
							>
								<div class="self-center p-1.5">
									<Sidebar />
								</div>
							</button>
						</Tooltip>
					</div>
				{/if}

				<div class="ml-2 py-0.5 self-center flex items-center justify-between w-full">
					<div class="flex gap-1 text-sm font-medium py-1">
						<span>{$i18n.t('FAQ')}</span>
						<span class="text-xs text-gray-500 dark:text-gray-400 self-center ml-2">(Admin only)</span>
					</div>
				</div>
			</div>
		</nav>

		<!-- Content -->
		<div class="flex-1 overflow-y-auto">
			<div class="max-w-3xl mx-auto px-4 py-6">
				<!-- Title -->
				<div class="mb-6">
					<h1 class="text-2xl font-semibold text-gray-800 dark:text-gray-100">
						{$i18n.t('Laila FAQ')}
					</h1>
					<p class="text-sm text-gray-500 dark:text-gray-400 mt-1">
						{$i18n.t('Find answers to common questions about using Laila.')}
					</p>
				</div>

				<!-- Search & Filters -->
				<div class="mb-6 space-y-3">
					<!-- Search -->
					<div
						class="flex items-center gap-2 px-3 py-2 bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-700"
					>
						<Search className="size-4 text-gray-400" />
						<input
							type="text"
							class="flex-1 text-sm bg-transparent outline-none placeholder-gray-400"
							placeholder={$i18n.t('Search FAQ...')}
							bind:value={query}
						/>
						{#if query}
							<button
								class="text-gray-400 hover:text-gray-600 dark:hover:text-gray-300"
								on:click={() => (query = '')}
							>
								<XMark className="size-4" />
							</button>
						{/if}
					</div>

					<!-- Category Filters -->
					<div class="flex flex-wrap gap-2">
						<button
							class="px-3 py-1.5 text-xs font-medium rounded-lg transition {selectedCategory === ''
								? 'bg-gray-200 dark:bg-gray-700 text-gray-800 dark:text-gray-100'
								: 'bg-gray-100 dark:bg-gray-800 text-gray-600 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-gray-700'}"
							on:click={() => (selectedCategory = '')}
						>
							{$i18n.t('All')}
						</button>
						{#each categories as category}
							<button
								class="px-3 py-1.5 text-xs font-medium rounded-lg transition {selectedCategory ===
								category.id
									? 'bg-gray-200 dark:bg-gray-700 text-gray-800 dark:text-gray-100'
									: 'bg-gray-100 dark:bg-gray-800 text-gray-600 dark:text-gray-400 hover:bg-gray-200 dark:hover:bg-gray-700'}"
								on:click={() =>
									(selectedCategory = selectedCategory === category.id ? '' : category.id)}
							>
								{$i18n.t(category.label)}
							</button>
						{/each}
					</div>
				</div>

				<!-- FAQ Items -->
				{#if filteredItems.length > 0}
					<div class="space-y-3">
						{#each filteredItems as item (item.id)}
							<details
								class="group bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-700 overflow-hidden"
							>
								<summary
									class="flex items-center justify-between px-4 py-3 cursor-pointer list-none hover:bg-gray-50 dark:hover:bg-gray-850 transition"
								>
									<span class="text-sm font-medium text-gray-800 dark:text-gray-100 pr-4">
										{item.question}
									</span>
									<ChevronDown
										className="size-4 text-gray-400 transition-transform group-open:rotate-180"
									/>
								</summary>
								<div class="px-4 pb-4 pt-1">
									<p class="text-sm text-gray-600 dark:text-gray-300 leading-relaxed whitespace-pre-line">
										{item.answer}
									</p>
									<div class="mt-3">
										<span
											class="inline-block px-2 py-0.5 text-xs rounded-md bg-gray-100 dark:bg-gray-800 text-gray-500 dark:text-gray-400"
										>
											{categories.find((c) => c.id === item.category)?.label || item.category}
										</span>
									</div>
								</div>
							</details>
						{/each}
					</div>
				{:else}
					<!-- Empty State -->
					<div class="flex flex-col items-center justify-center py-16 text-center">
						<div class="text-4xl mb-3">?</div>
						<h3 class="text-lg font-medium text-gray-700 dark:text-gray-300 mb-1">
							{$i18n.t('No results found')}
						</h3>
						<p class="text-sm text-gray-500 dark:text-gray-400">
							{$i18n.t('Try adjusting your search or filter to find what you are looking for.')}
						</p>
					</div>
				{/if}

				<!-- Contact Section -->
				<div class="mt-8 p-4 bg-gray-50 dark:bg-gray-850 rounded-xl">
					<h3 class="text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">
						{$i18n.t("Can't find what you're looking for?")}
					</h3>
					<p class="text-sm text-gray-500 dark:text-gray-400">
						{$i18n.t('Contact your administrator or reach out to support for further assistance.')}
					</p>
				</div>
			</div>
		</div>
	</div>
{/if}
