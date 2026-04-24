<script lang="ts">
  export let data;
</script>

<svelte:head>
  <title>{data.doc.title} | {data.locale.siteTitle}</title>
</svelte:head>

<article class="panel doc-body">
  <h2 class="doc-title">{data.doc.title}</h2>

  <div class="doc-html">
    {@html data.doc.html}
  </div>

  <div class="footer-note">
    <p class="muted">{data.locale.generatedNotice}</p>
    <p class="meta">
      {data.locale.sourceLabel}: <code>{data.doc.sourcePath}</code>
    </p>
  </div>

  <a class="button-link" href={`/${data.lang}/docs`}>{data.locale.backToDocs}</a>
</article>
