declare global {
  namespace App {
    interface PageData {
      lang?: 'cs' | 'en';
      locale?: import('$lib/content/site').LocaleContent;
    }
  }
}

export {};

