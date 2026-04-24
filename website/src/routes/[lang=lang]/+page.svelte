<script lang="ts">
  import { installersByLanguage } from '$lib/generated/content';

  export let data;

  const installers = [...installersByLanguage.cs, ...installersByLanguage.en];
</script>

<svelte:head>
  <title>{data.locale.siteTitle} | {data.locale.navHome}</title>
</svelte:head>

<section class="hero hero-home">
  <div class="hero-copy">
    <div class="eyebrow">Excel Statistics Add-in</div>
    <h2>{data.locale.heroTitle}</h2>
    <p>{data.locale.heroBody}</p>
  </div>

  <div class="hero-links">
    <a class="link-card link-card-primary" href={`/${data.lang}/docs`}>
      <div class="eyebrow">{data.locale.navDocs}</div>
      <h3 class="section-title">{data.locale.docsTitle}</h3>
      <p class="lede">{data.locale.docsBody}</p>
    </a>

    <article class="panel panel-compact">
      <div class="eyebrow">{data.locale.navDownloads}</div>
      <h3 class="section-title">{data.locale.downloadsTitle}</h3>
      <p class="lede">{data.locale.downloadsBody}</p>
      <div class="card-list">
        {#each installers as installer}
          <div class="download-card">
            <strong>
              {installer.lang === 'cs' ? data.locale.installerCz : data.locale.installerEn}
            </strong>
            {#if installer.exists}
              <a class="button-link" href={installer.href}>{installer.fileName}</a>
            {:else}
              <span class="muted">{data.locale.installerMissing}</span>
            {/if}
          </div>
        {/each}
      </div>
    </article>
  </div>
</section>
