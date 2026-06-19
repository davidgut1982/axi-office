---
"@axi-office/excel": patch
"@axi-office/outlook": patch
"@axi-office/word": patch
"@axi-office/core": patch
---

Security remediations for public release: pin upstream MCP server versions (supply-chain TOFU fix), add opt-in path sandbox to word commands (--base-dir), harden outlook README with Entra app registration and token security guidance, mark callHttpTool @internal with SSRF note, add disclaimer to all READMEs.
