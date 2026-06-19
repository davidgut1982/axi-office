/**
 * Why: Markdown is the most common authoring format for agents; `from-markdown` converts a
 * .md file into a .docx without a heavy toolchain.
 * What: Validates <in.md> <out.docx>, parses the markdown into a DocSpec, builds and writes
 * the docx.
 * Test: Write a temp .md with "# Title", run fromMarkdownCommand([md, out]), assert the
 * output .docx exists and mammoth extracts the heading text.
 */
import { readFileSync, writeFileSync } from "node:fs";
import { AxiError, parseFlags } from "@axi-office/core";
import { buildDocxBuffer } from "../docx-build.js";
import { parseMarkdown } from "../markdown.js";

export async function fromMarkdownCommand(args: string[]): Promise<unknown> {
  const { positionals } = parseFlags(args);
  const [input, out] = positionals;
  if (!input || !out) {
    throw new AxiError("input.md and out.docx are required", "VALIDATION_ERROR", [
      "word-axi from-markdown <in.md> <out.docx>",
    ]);
  }

  let md: string;
  try {
    md = readFileSync(input, "utf8");
  } catch {
    throw new AxiError(`cannot read markdown file: ${input}`, "IO_ERROR");
  }

  const spec = parseMarkdown(md);
  const buffer = await buildDocxBuffer(spec);
  writeFileSync(out, buffer);

  return { created: out, bytes: buffer.length, sections: spec.sections.length };
}
