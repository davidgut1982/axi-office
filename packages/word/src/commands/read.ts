import { AxiError, parseFlags } from "@axi-office/core";
/**
 * Why: Agents need to read existing .docx content as plain text or HTML for downstream
 * processing; `read` exposes mammoth's extraction.
 * What: Validates <in.docx>, then returns raw text (default) or HTML based on --format.
 * Test: Create a docx with known text, run readCommand([file]); assert the returned text
 * contains that text.
 */
import mammoth from "mammoth";

export async function readCommand(args: string[]): Promise<unknown> {
  const { positionals, flags } = parseFlags(args);
  const file = positionals[0];
  if (!file) {
    throw new AxiError("input.docx is required", "VALIDATION_ERROR", [
      "word-axi read <in.docx> [--format raw|html]",
    ]);
  }

  const format = typeof flags.format === "string" ? flags.format : "raw";
  if (format !== "raw" && format !== "html") {
    throw new AxiError("--format must be raw or html", "VALIDATION_ERROR");
  }

  if (format === "html") {
    const result = await mammoth.convertToHtml({ path: file });
    return { file, format, html: result.value };
  }

  const result = await mammoth.extractRawText({ path: file });
  return { file, format, text: result.value };
}
