import { localeByLanguage, type LanguageCode } from '$lib/content/site';

export const prerender = true;

export function load({ params }) {
  const lang = params.lang as LanguageCode;

  return {
    lang,
    locale: localeByLanguage[lang]
  };
}

