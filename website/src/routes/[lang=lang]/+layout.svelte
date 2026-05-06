<script lang="ts">
  import { goto } from '$app/navigation';
  import { page } from '$app/stores';

  export let data;

  function isActive(target: string) {
    return $page.url.pathname === target;
  }

  function switchPath(targetLang: 'cs' | 'en') {
    return `/${targetLang}${$page.url.pathname.replace(/^\/(cs|en)/, '') || ''}`;
  }

  async function switchLanguage(event: MouseEvent, targetLang: 'cs' | 'en') {
    event.preventDefault();
    await goto(switchPath(targetLang));
  }
</script>

<svelte:head>
  <title>{data.locale.siteTitle}</title>
</svelte:head>

<div class="shell">
  <div class="frame">
    <header class="topbar">
      <div class="brand">
        <h1>{data.locale.siteTitle}</h1>
        <p>{data.locale.siteTagline}</p>
      </div>

      <nav class="nav" aria-label="Primary">
        <a class:active={isActive(`/${data.lang}`)} href={`/${data.lang}`}>{data.locale.navHome}</a>
        <a class:active={$page.url.pathname.startsWith(`/${data.lang}/docs`)} href={`/${data.lang}/docs`}>
          {data.locale.navDocs}
        </a>
        <a class:active={$page.url.pathname.startsWith(`/${data.lang}/downloads`)} href={`/${data.lang}/downloads`}>
          {data.locale.navDownloads}
        </a>
      </nav>

      <div class="lang-switch">
        <a
          class:active={data.lang === 'cs'}
          href={switchPath('cs')}
          on:click={(event) => switchLanguage(event, 'cs')}
        >
          CS
        </a>
        <a
          class:active={data.lang === 'en'}
          href={switchPath('en')}
          on:click={(event) => switchLanguage(event, 'en')}
        >
          EN
        </a>
      </div>
    </header>

    <main class="content">
      <slot />
    </main>
  </div>
</div>
