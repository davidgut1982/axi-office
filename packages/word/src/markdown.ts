/**
 * Why: from-markdown needs to turn Markdown into the shared DocSpec without pulling in a
 * heavy Markdown library; a small line-based parser covers the common subset (headings,
 * paragraphs, bullet lists, bold/italic).
 * What: parseMarkdown(md) returns a DocSpec; parseInline(text) splits a line into
 * bold/italic runs.
 * Test: parseMarkdown("# Title\\n\\ntext") yields a heading section then a paragraph;
 * parseInline("a **b** c") yields three runs with the middle one bold.
 */
import type { DocSection, DocSpec, InlineRun } from "./docx-build.js";

/**
 * Why: Inline emphasis (**bold**, *italic*) must become separate docx runs.
 * What: Tokenizes a single line into InlineRun[] honoring ** and * delimiters.
 * Test: parseInline("**x**") returns [{ text: "x", bold: true }].
 */
export function parseInline(text: string): InlineRun[] {
  const runs: InlineRun[] = [];
  const regex = /(\*\*([^*]+)\*\*)|(\*([^*]+)\*)/g;
  let lastIndex = 0;
  let match: RegExpExecArray | null = regex.exec(text);

  while (match !== null) {
    if (match.index > lastIndex) {
      runs.push({ text: text.slice(lastIndex, match.index) });
    }
    if (match[2] !== undefined) {
      runs.push({ text: match[2], bold: true });
    } else if (match[4] !== undefined) {
      runs.push({ text: match[4], italics: true });
    }
    lastIndex = regex.lastIndex;
    match = regex.exec(text);
  }

  if (lastIndex < text.length) {
    runs.push({ text: text.slice(lastIndex) });
  }
  return runs.length > 0 ? runs : [{ text }];
}

/**
 * Why: Produces the shared DocSpec from Markdown so buildDocxBuffer can render it.
 * What: Parses headings (#, ##, ###...), bullet lists (- / *), and paragraphs; blank
 * lines separate blocks. Consecutive bullet lines are grouped into one list section.
 * Test: parseMarkdown("- a\\n- b") yields one list section with items ["a","b"].
 */
export function parseMarkdown(md: string): DocSpec {
  const sections: DocSection[] = [];
  const lines = md.split(/\r?\n/);
  let listBuffer: string[] = [];

  const flushList = () => {
    if (listBuffer.length > 0) {
      sections.push({ type: "list", items: listBuffer });
      listBuffer = [];
    }
  };

  for (const rawLine of lines) {
    const line = rawLine.trimEnd();
    if (line.trim() === "") {
      flushList();
      continue;
    }

    const heading = /^(#{1,6})\s+(.*)$/.exec(line);
    if (heading) {
      flushList();
      sections.push({
        type: "heading",
        level: heading[1].length,
        text: heading[2].trim(),
      });
      continue;
    }

    const bullet = /^[-*]\s+(.*)$/.exec(line.trim());
    if (bullet) {
      listBuffer.push(bullet[1].trim());
      continue;
    }

    flushList();
    sections.push({ type: "paragraph", runs: parseInline(line.trim()) });
  }

  flushList();
  return { sections };
}
