export type LanguageCode = 'cs' | 'en';

export type LocaleContent = {
  code: LanguageCode;
  siteTitle: string;
  siteTagline: string;
  languageLabel: string;
  otherLanguageLabel: string;
  navHome: string;
  navDocs: string;
  navDownloads: string;
  heroTitle: string;
  heroBody: string;
  docsTitle: string;
  docsBody: string;
  docsSearchPlaceholder: string;
  downloadsTitle: string;
  downloadsBody: string;
  docsEmpty: string;
  backToDocs: string;
  sourceLabel: string;
  installerCz: string;
  installerEn: string;
  installerMissing: string;
  generatedNotice: string;
};

export const localeByLanguage: Record<LanguageCode, LocaleContent> = {
  cs: {
    code: 'cs',
    siteTitle: 'XLStatUDF',
    siteTagline: 'Statistický doplněk pro Microsoft Excel',
    languageLabel: 'Čeština',
    otherLanguageLabel: 'English',
    navHome: 'Úvod',
    navDocs: 'Dokumentace',
    navDownloads: 'Stažení',
    heroTitle: 'XLStatUDF pro Microsoft Excel',
    heroBody:
      'XLStatUDF je add-in pro Excel, který pomocí vlastních funkcí překlenuje mezeru mezi základním Excelem a statistickými softwary. Jeho cílem je zjednodušit výuku statistiky, kdy není hlavním zájmem učit studenty převádět vzorce do Excelu, ale naučit je používat statistiku k řešení problémů.',
    docsTitle: 'Rejstřík funkcí',
    docsBody:
      'Vyhledávatelný přehled dostupných excelových funkcí s odkazy na jejich detailní dokumentaci.',
    docsSearchPlaceholder: 'Hledat název funkce nebo popis...',
    downloadsTitle: 'Stažení',
    downloadsBody:
      'Dostupné instalační soubory pro jednotlivé jazykové verze doplňku.',
    docsEmpty: 'Nenalezena žádná odpovídající funkce.',
    backToDocs: 'Zpět na dokumentaci',
    sourceLabel: 'Zdroj v repozitáři',
    installerCz: 'Český instalátor',
    installerEn: 'English installer',
    installerMissing: 'Instalační soubor zatím není k dispozici.',
    generatedNotice: 'Obsah této stránky je generovaný z repozitářové dokumentace.'
  },
  en: {
    code: 'en',
    siteTitle: 'XLStatUDF',
    siteTagline: 'Statistical add-in for Microsoft Excel',
    languageLabel: 'English',
    otherLanguageLabel: 'Čeština',
    navHome: 'Home',
    navDocs: 'Documentation',
    navDownloads: 'Downloads',
    heroTitle: 'XLStatUDF for Microsoft Excel',
    heroBody:
      'XLStatUDF is an Excel add-in that bridges the gap between standard spreadsheets and dedicated statistical software through custom worksheet functions. Its purpose is to simplify statistics teaching, where the main goal is not to train students to translate formulas into Excel, but to teach them how to use statistics to solve problems.',
    docsTitle: 'Function Index',
    docsBody:
      'Searchable index of available Excel functions with links to detailed documentation.',
    docsSearchPlaceholder: 'Search by function name or description...',
    downloadsTitle: 'Downloads',
    downloadsBody:
      'Available installer packages for the currently published language builds.',
    docsEmpty: 'No matching function was found.',
    backToDocs: 'Back to documentation',
    sourceLabel: 'Repository source',
    installerCz: 'Czech installer',
    installerEn: 'English installer',
    installerMissing: 'Installer file is not available yet.',
    generatedNotice: 'This page is generated from repository documentation.'
  }
};
