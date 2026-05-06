import { mkdir, readdir, readFile, writeFile, copyFile, stat } from 'node:fs/promises';
import path from 'node:path';

const websiteRoot = path.resolve(import.meta.dirname, '..');
const repoRoot = path.resolve(websiteRoot, '..');
const docsRoot = path.join(repoRoot, 'docs');
const generatedFile = path.join(websiteRoot, 'src', 'lib', 'generated', 'content.ts');
const downloadDir = path.join(websiteRoot, 'static', 'downloads');

const languageLabels = {
  cs: 'Czech',
  en: 'English'
};

const installerDefinitions = [
  { lang: 'cs', label: 'Evalytics CS Setup', fileName: 'Evalytics_CS_Setup.exe' },
  { lang: 'en', label: 'Evalytics EN Setup', fileName: 'Evalytics_EN_Setup.exe' }
];

async function listMarkdownFiles(root, relative = '') {
  const absolute = path.join(root, relative);
  const entries = await readdir(absolute, { withFileTypes: true });
  const results = [];

  for (const entry of entries) {
    const nextRelative = path.join(relative, entry.name);
    if (entry.isDirectory()) {
      results.push(...(await listMarkdownFiles(root, nextRelative)));
      continue;
    }

    if (entry.isFile() && entry.name.endsWith('.md')) {
      results.push(nextRelative);
    }
  }

  return results.sort((left, right) => left.localeCompare(right));
}

function toPosix(value) {
  return value.split(path.sep).join('/');
}

function sectionName(relativePath) {
  const normalized = toPosix(relativePath);
  if (normalized.startsWith('functions/')) {
    return 'Functions';
  }

  if (normalized === 'README.md') {
    return 'Overview';
  }

  return normalized.split('/')[0].replace(/\.md$/, '');
}

function slugify(value) {
  return stripMarkdown(value)
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9_\s-]+/g, '')
    .replace(/\s+/g, '-')
    .replace(/^-+|-+$/g, '');
}

function slugFromPath(relativePath) {
  const normalized = toPosix(relativePath);
  if (normalized === 'README.md') {
    return 'overview';
  }

  return normalized.replace(/\.md$/, '').replace(/\/README$/i, '');
}

function stripMarkdown(markdown) {
  return markdown
    .replace(/```[\s\S]*?```/g, ' ')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/\[([^\]]+)\]\([^)]+\)/g, '$1')
    .replace(/^#+\s+/gm, '')
    .replace(/^\|.*\|$/gm, ' ')
    .replace(/[*_>#-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function escapeHtml(value) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function rewriteHref(href, lang, relativePath) {
  if (href.startsWith('http://') || href.startsWith('https://') || href.startsWith('#')) {
    return href;
  }

  const installerMatch = href.match(/(?:Evalytics|XLStatUDF)_(CS|EN)_Setup\.exe/i);
  if (installerMatch) {
    return `/downloads/${installerMatch[0]}`;
  }

  if (/artifacts\/main\/.+\.xll$/i.test(href) || /\/reference\//i.test(href)) {
    return '#';
  }

  if (/^\/?[a-z]:[\\/]/i.test(href)) {
    return '#';
  }

  const docsAbsoluteMatch = href.match(/docs\/(cs|en)\/(.+?)\.md$/i);
  if (docsAbsoluteMatch) {
    const [, absoluteLang, docPath] = docsAbsoluteMatch;
    const slug = slugFromPath(docPath);
    return `/${absoluteLang.toLowerCase()}/docs/${slug}`;
  }

  const markdownLinkMatch = href.match(/^(.+?\.md)(#(.+))?$/i);
  if (markdownLinkMatch) {
    const [, markdownHref, , anchorText] = markdownLinkMatch;
    const currentDir = path.posix.dirname(toPosix(relativePath));
    const resolved = path.posix.normalize(path.posix.join(currentDir, markdownHref));
    const slug = slugFromPath(resolved);
    const anchor = anchorText ? `#${slugify(anchorText)}` : '';
    return `/${lang}/docs/${slug}${anchor}`;
  }

  return href;
}

function inlineMarkdown(value, lang, relativePath) {
  return escapeHtml(value)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, (_match, text, href) => {
      const safeHref = escapeHtml(rewriteHref(href, lang, relativePath));
      return `<a href="${safeHref}">${text}</a>`;
    });
}

function isTableSeparator(line) {
  const trimmed = line.trim();
  return /^\|(?:\s*:?-+:?\s*\|)+$/.test(trimmed);
}

function markdownToHtml(markdown, lang, relativePath) {
  const lines = markdown.replace(/\r\n/g, '\n').split('\n');
  const html = [];
  let index = 0;

  while (index < lines.length) {
    const line = lines[index];
    const trimmed = line.trim();

    if (!trimmed) {
      index += 1;
      continue;
    }

    if (trimmed.startsWith('```')) {
      const language = trimmed.slice(3).trim();
      const block = [];
      index += 1;

      while (index < lines.length && !lines[index].trim().startsWith('```')) {
        block.push(lines[index]);
        index += 1;
      }

      index += 1;
      const className = language ? ` class="language-${escapeHtml(language)}"` : '';
      html.push(`<pre><code${className}>${escapeHtml(block.join('\n'))}</code></pre>`);
      continue;
    }

    if (/^#{1,6}\s+/.test(trimmed)) {
      const level = trimmed.match(/^#+/)[0].length;
      const content = trimmed.replace(/^#{1,6}\s+/, '');
      const headingId = slugify(content);
      const idAttribute = headingId ? ` id="${headingId}"` : '';
      html.push(`<h${level}${idAttribute}>${inlineMarkdown(content, lang, relativePath)}</h${level}>`);
      index += 1;
      continue;
    }

    if (trimmed.startsWith('- ')) {
      const items = [];
      while (index < lines.length && lines[index].trim().startsWith('- ')) {
        items.push(lines[index].trim().slice(2));
        index += 1;
      }

      html.push(
        `<ul>${items.map((item) => `<li>${inlineMarkdown(item, lang, relativePath)}</li>`).join('')}</ul>`
      );
      continue;
    }

    if (trimmed.startsWith('|') && index + 1 < lines.length && isTableSeparator(lines[index + 1])) {
      const header = trimmed
        .split('|')
        .slice(1, -1)
        .map((cell) => cell.trim());
      index += 2;

      const rows = [];
      while (index < lines.length && lines[index].trim().startsWith('|')) {
        rows.push(
          lines[index]
            .trim()
            .split('|')
            .slice(1, -1)
            .map((cell) => cell.trim())
        );
        index += 1;
      }

      const thead = `<thead><tr>${header
        .map((cell) => `<th>${inlineMarkdown(cell, lang, relativePath)}</th>`)
        .join('')}</tr></thead>`;
      const tbody = `<tbody>${rows
        .map(
          (row) =>
            `<tr>${row
              .map((cell) => `<td>${inlineMarkdown(cell, lang, relativePath)}</td>`)
              .join('')}</tr>`
        )
        .join('')}</tbody>`;
      html.push(`<table>${thead}${tbody}</table>`);
      continue;
    }

    const paragraph = [trimmed];
    index += 1;

    while (index < lines.length) {
      const candidate = lines[index].trim();
      if (
        !candidate ||
        candidate.startsWith('```') ||
        /^#{1,6}\s+/.test(candidate) ||
        candidate.startsWith('- ') ||
        (candidate.startsWith('|') && index + 1 < lines.length && isTableSeparator(lines[index + 1]))
      ) {
        break;
      }

      paragraph.push(candidate);
      index += 1;
    }

    html.push(`<p>${inlineMarkdown(paragraph.join(' '), lang, relativePath)}</p>`);
  }

  return html.join('\n');
}

async function buildDocs(lang) {
  const root = path.join(docsRoot, lang);
  const files = await listMarkdownFiles(root);
  const docs = [];

  for (const relativePath of files) {
    const absolutePath = path.join(root, relativePath);
    const markdown = await readFile(absolutePath, 'utf8');
    const lines = markdown.replace(/\r\n/g, '\n').split('\n');
    const titleLine = lines.find((line) => line.trim().startsWith('# ')) ?? path.basename(relativePath, '.md');
    const title = titleLine.replace(/^#\s+/, '').trim();
    const plain = stripMarkdown(markdown);
    const summary = plain.slice(0, 180).trim();

    docs.push({
      lang,
      slug: slugFromPath(relativePath),
      section: sectionName(relativePath),
      title,
      summary,
      sourcePath: toPosix(path.relative(repoRoot, absolutePath)),
      html: markdownToHtml(markdown, lang, relativePath)
    });
  }

  return docs;
}

function buildFunctionHref(lang, relativeHref) {
  const [filePart, anchorPart] = relativeHref.split('#');
  const normalizedPath = filePart.replace(/^\.\//, '');
  const resolvedPath = normalizedPath.startsWith('functions/')
    ? normalizedPath
    : path.posix.normalize(path.posix.join('functions', normalizedPath));
  const slug = slugFromPath(resolvedPath);
  const anchor = anchorPart ? `#${slugify(anchorPart)}` : '';
  return `/${lang}/docs/${slug}${anchor}`;
}

async function buildFunctionIndex(lang) {
  const indexPath = path.join(docsRoot, lang, 'functions', 'index.md');
  const markdown = await readFile(indexPath, 'utf8');
  const lines = markdown.replace(/\r\n/g, '\n').split('\n');
  const entries = [];

  for (const line of lines) {
    const match = line.match(/^- \[([^\]]+)\]\(([^)]+)\):\s*(.+)$/);
    if (!match) {
      continue;
    }

    const [, name, href, summary] = match;
    entries.push({
      lang,
      name: name.trim(),
      summary: summary.trim(),
      href: buildFunctionHref(lang, href.trim())
    });
  }

  return entries;
}

async function fileExists(target) {
  try {
    const info = await stat(target);
    return info.isFile();
  } catch {
    return false;
  }
}

async function syncInstallers() {
  await mkdir(downloadDir, { recursive: true });

  const installers = { cs: [], en: [] };

  for (const definition of installerDefinitions) {
    const source = path.join(repoRoot, 'artifacts', 'installer', definition.fileName);
    const target = path.join(downloadDir, definition.fileName);
    const exists = await fileExists(source);

    if (exists) {
      await copyFile(source, target);
    }

    installers[definition.lang].push({
      lang: definition.lang,
      label: `${languageLabels[definition.lang]} installer`,
      fileName: definition.fileName,
      href: `/downloads/${definition.fileName}`,
      exists
    });
  }

  return installers;
}

function serialize(value) {
  return JSON.stringify(value, null, 2);
}

async function main() {
  const [docsCs, docsEn, functionsCs, functionsEn, installers] = await Promise.all([
    buildDocs('cs'),
    buildDocs('en'),
    buildFunctionIndex('cs'),
    buildFunctionIndex('en'),
    syncInstallers()
  ]);

  const generatedAt = new Date().toISOString();
  const content = `export type GeneratedDoc = {
  lang: 'cs' | 'en';
  slug: string;
  section: string;
  title: string;
  summary: string;
  sourcePath: string;
  html: string;
};

export type GeneratedInstaller = {
  lang: 'cs' | 'en';
  label: string;
  fileName: string;
  href: string;
  exists: boolean;
};

export type FunctionIndexEntry = {
  lang: 'cs' | 'en';
  name: string;
  summary: string;
  href: string;
};

export const generatedAt = ${serialize(generatedAt)};
export const docsByLanguage: Record<'cs' | 'en', GeneratedDoc[]> = ${serialize({
    cs: docsCs,
    en: docsEn
  })};

export const functionIndexByLanguage: Record<'cs' | 'en', FunctionIndexEntry[]> = ${serialize({
    cs: functionsCs,
    en: functionsEn
  })};

export const installersByLanguage: Record<'cs' | 'en', GeneratedInstaller[]> = ${serialize(installers)};
`;

  await mkdir(path.dirname(generatedFile), { recursive: true });
  await writeFile(generatedFile, content, 'utf8');
  console.log(`Synchronized ${docsCs.length + docsEn.length} documentation pages.`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
