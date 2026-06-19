/**
 * Why: Both `create` (JSON spec) and `from-markdown` need to turn an in-memory document
 * model into docx elements; sharing one builder avoids duplicating the docx element
 * construction logic and the heading-level mapping.
 * What: Defines a small DocSpec model and buildDocxBuffer(spec) which constructs a docx
 * Document and returns the packed .docx Buffer.
 * Test: Call buildDocxBuffer({ sections: [{ type: "heading", level: 1, text: "Hi" }] }),
 * assert the result is a non-empty Buffer; an integration test round-trips it via mammoth.
 */
import {
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
} from "docx";

export type DocSection =
  | { type: "heading"; level?: number; text: string }
  | { type: "paragraph"; text?: string; runs?: InlineRun[] }
  | { type: "list"; items: string[] }
  | { type: "table"; rows: string[][] };

export interface InlineRun {
  text: string;
  bold?: boolean;
  italics?: boolean;
}

export interface DocSpec {
  title?: string;
  sections: DocSection[];
}

const HEADING_BY_LEVEL: Record<number, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
};

/**
 * Why: docx requires a typed HeadingLevel; specs use numeric 1-6, so we map and clamp.
 * What: Returns the HeadingLevel for a 1-6 number, defaulting to HEADING_1.
 * Test: headingFor(2) === HeadingLevel.HEADING_2; headingFor(99) === HeadingLevel.HEADING_1.
 */
function headingFor(level: number | undefined): (typeof HeadingLevel)[keyof typeof HeadingLevel] {
  return HEADING_BY_LEVEL[level ?? 1] ?? HeadingLevel.HEADING_1;
}

/**
 * Why: A list/table/paragraph section maps to one or more docx Paragraph/Table children;
 * isolating the mapping keeps buildDocxBuffer readable.
 * What: Converts a single DocSection into an array of docx children.
 * Test: sectionToChildren({ type: "list", items: ["a"] }) returns one Paragraph with
 * bullet numbering.
 */
function sectionToChildren(section: DocSection): (Paragraph | Table)[] {
  switch (section.type) {
    case "heading":
      return [
        new Paragraph({
          heading: headingFor(section.level),
          text: section.text,
        }),
      ];
    case "paragraph": {
      if (section.runs && section.runs.length > 0) {
        return [
          new Paragraph({
            children: section.runs.map(
              (r) => new TextRun({ text: r.text, bold: r.bold, italics: r.italics })
            ),
          }),
        ];
      }
      return [new Paragraph({ text: section.text ?? "" })];
    }
    case "list":
      return section.items.map((item) => new Paragraph({ text: item, bullet: { level: 0 } }));
    case "table":
      return [
        new Table({
          rows: section.rows.map(
            (row) =>
              new TableRow({
                children: row.map(
                  (cell) =>
                    new TableCell({
                      children: [new Paragraph({ text: cell })],
                    })
                ),
              })
          ),
        }),
      ];
    default: {
      const _exhaustive: never = section;
      throw new Error(`Unknown section type: ${JSON.stringify(_exhaustive)}`);
    }
  }
}

/**
 * Why: Callers need a ready-to-write .docx buffer from a spec without touching docx
 * internals.
 * What: Builds a Document with an optional title heading plus all sections, then packs it.
 * Test: Pass a spec with a title and a paragraph; assert the returned Buffer length > 0.
 */
export async function buildDocxBuffer(spec: DocSpec): Promise<Buffer> {
  const children: (Paragraph | Table)[] = [];
  if (spec.title) {
    children.push(new Paragraph({ heading: HeadingLevel.TITLE, text: spec.title }));
  }
  for (const section of spec.sections) {
    children.push(...sectionToChildren(section));
  }

  const doc = new Document({ sections: [{ children }] });
  return Packer.toBuffer(doc);
}
