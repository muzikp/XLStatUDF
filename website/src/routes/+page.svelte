<script lang="ts">
  import { goto } from '$app/navigation';
  import { onMount } from 'svelte';

  function resolveLanguage() {
    if (typeof window === 'undefined') {
      return 'en';
    }

    return window.navigator.language.toLowerCase().startsWith('cs') ? 'cs' : 'en';
  }

  onMount(async () => {
    const lang = resolveLanguage();
    await goto(`/${lang}`, { replaceState: true });
  });
</script>

<svelte:head>
  <title>XLStatUDF</title>
  <meta
    name="description"
    content="Static bilingual documentation and download site for XLStatUDF."
  />
</svelte:head>

<div class="shell">
  <div class="frame">
    <div class="content">
      <section class="hero">
        <div class="eyebrow">XLStatUDF</div>
        <h2>Redirecting to your preferred language</h2>
        <p>
          The site remembers your language choice and sends you directly to the matching version.
        </p>
      </section>
    </div>
  </div>
</div>
