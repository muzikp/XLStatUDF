import { error } from '@sveltejs/kit';
import { docsByLanguage } from '$lib/generated/content';
import type { LanguageCode } from '$lib/content/site';

export const prerender = true;

export function entries() {
  return [...docsByLanguage.cs, ...docsByLanguage.en].map((doc) => ({
    lang: doc.lang,
    slug: doc.slug.split('/')
  }));
}

export function load({ params }) {
  const lang = params.lang as LanguageCode;
  const slug = params.slug;
  const doc = docsByLanguage[lang].find((entry) => entry.slug === slug);

  if (!doc) {
    throw error(404, 'Document not found');
  }

  return { doc };
}

