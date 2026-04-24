import { docsByLanguage } from '$lib/generated/content';
import type { LanguageCode } from '$lib/content/site';

export function getDocs(lang: LanguageCode) {
  return docsByLanguage[lang];
}

export function getDoc(lang: LanguageCode, slug: string) {
  return docsByLanguage[lang].find((doc) => doc.slug === slug) ?? null;
}

