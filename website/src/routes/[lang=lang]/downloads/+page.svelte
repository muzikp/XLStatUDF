<script lang="ts">
  import { installersByLanguage } from '$lib/generated/content';

  export let data;

  const installers = [...installersByLanguage.cs, ...installersByLanguage.en];
</script>

<svelte:head>
  <title>{data.locale.siteTitle} | {data.locale.navDownloads}</title>
</svelte:head>

<section class="hero">
  <div class="eyebrow">{data.locale.navDownloads}</div>
  <h2>{data.locale.downloadsTitle}</h2>
  <p>{data.locale.downloadsBody}</p>
</section>

<section class="panel">
  <div class="card-list">
    {#each installers as installer}
      <div class="download-card">
        <strong>{installer.lang === 'cs' ? data.locale.installerCz : data.locale.installerEn}</strong>
        <span class="meta">{installer.fileName}</span>
        {#if installer.exists}
          <a class="button-link" href={installer.href}>{installer.fileName}</a>
        {:else}
          <span class="muted">{data.locale.installerMissing}</span>
        {/if}
      </div>
    {/each}
  </div>
</section>
