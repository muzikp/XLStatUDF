<script lang="ts">
  import { functionIndexByLanguage } from '$lib/generated/content';

  export let data;

  let query = '';

  $: entries = functionIndexByLanguage[data.lang];
  $: normalizedQuery = query.trim().toLowerCase();
  $: filteredEntries =
    normalizedQuery.length === 0
      ? entries
      : entries.filter((entry) =>
          `${entry.name} ${entry.summary}`.toLowerCase().includes(normalizedQuery)
        );
</script>

<svelte:head>
  <title>{data.locale.siteTitle} | {data.locale.navDocs}</title>
</svelte:head>

<section class="hero hero-index">
  <div class="eyebrow">{data.locale.navDocs}</div>
  <h2>{data.locale.docsTitle}</h2>
  <p>{data.locale.docsBody}</p>
</section>

<section class="panel">
  <div class="search-shell">
    <input
      bind:value={query}
      class="search-input"
      type="search"
      placeholder={data.locale.docsSearchPlaceholder}
      aria-label={data.locale.docsSearchPlaceholder}
    />
  </div>

  <div class="function-index">
    {#if filteredEntries.length === 0}
      <p class="muted">{data.locale.docsEmpty}</p>
    {:else}
      {#each filteredEntries as entry}
        <a class="function-row" href={entry.href}>
          <div class="function-formula">={entry.name}</div>
          <div class="function-summary"><em>{entry.summary}</em></div>
        </a>
      {/each}
    {/if}
  </div>
</section>
