#!/usr/bin/env node
/**
 * Why: This machine uses a ZTK shell wrapper that intercepts certain process
 * spawning patterns. By using spawnSync with shell:false and an explicit env,
 * we bypass the wrapper and invoke biome directly.
 */
import { spawnSync } from "node:child_process";
import { join } from "node:path";
import { existsSync } from "node:fs";

const pkgDir = process.cwd();
const src = join(pkgDir, "src");
const test = join(pkgDir, "test");

const localBiome = join(pkgDir, "node_modules", "@biomejs", "biome", "bin", "biome");
const rootBiome = join(pkgDir, "..", "..", "node_modules", "@biomejs", "biome", "bin", "biome");
const biomeBin = existsSync(localBiome) ? localBiome : rootBiome;

const writeMode = process.argv.slice(2).some((a) => a === "--write" || a === "--fix");
const biomeArgs = writeMode
  ? ["check", "--write", src, test]
  : ["check", src, test];

const result = spawnSync(process.execPath, [biomeBin, ...biomeArgs], {
  stdio: "inherit",
  shell: false,
  env: process.env,
});

process.exit(result.status ?? 1);
