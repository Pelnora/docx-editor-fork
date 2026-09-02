#!/usr/bin/env node
/**
 * Pelnora fork packaging (H0-3).
 *
 * Builds nothing — run `bun run build` in packages/core and packages/react
 * first. Stages each package in a temp dir, rewrites the core's `name` to
 * `pelnora-docx-editor-core` (the source tree keeps `@eigenpal/docx-editor-core`
 * so workspace links, imports and the api-extractor scripts stay untouched),
 * runs `npm pack` there and drops the tarballs in the output directory.
 *
 * The react package is packed as-is: its dependency stays
 * `@eigenpal/docx-editor-core`, which the Pelnora app maps with one pnpm
 * override to `npm:pelnora-docx-editor-core@<version>` (the app depends on the
 * core directly too, so the override is needed either way).
 *
 * Usage: node scripts/pelnora/pack.mjs [outDir]   (default /tmp/pelnora-audit/fork-pack)
 * Publish (orchestrator, with OTP), core first:
 *   npm publish <outDir>/pelnora-docx-editor-core-<v>.tgz --access public --tag pelnora --otp=<code>
 *   npm publish <outDir>/pelnora-docx-editor-react-<v>.tgz --access public --tag pelnora --otp=<code>
 */

import {
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const outDir = resolve(process.argv[2] ?? '/tmp/pelnora-audit/fork-pack');
mkdirSync(outDir, { recursive: true });

const CORE_PUBLISH_NAME = 'pelnora-docx-editor-core';

function stageAndPack(pkgDir, { rename, requireFiles = [] } = {}) {
  const pkg = JSON.parse(readFileSync(join(pkgDir, 'package.json'), 'utf8'));
  for (const f of ['dist', ...requireFiles]) {
    if (!existsSync(join(pkgDir, f))) {
      throw new Error(`${pkg.name}: ${f} missing — run \`bun run build\` in ${pkgDir} first`);
    }
  }
  const staging = mkdtempSync(join(tmpdir(), `pelnora-pack-${basename(pkgDir)}-`));
  try {
    const files = Array.isArray(pkg.files) && pkg.files.length > 0 ? pkg.files : ['dist'];
    for (const f of new Set([...files, 'README.md', 'LICENSE', 'NOTICE', 'CHANGELOG.md'])) {
      if (existsSync(join(pkgDir, f)))
        cpSync(join(pkgDir, f), join(staging, f), { recursive: true });
    }
    if (rename) {
      pkg.name = rename;
    }
    // Lifecycle scripts are the monorepo's business, not the tarball's.
    delete pkg.scripts;
    delete pkg.devDependencies;
    writeFileSync(join(staging, 'package.json'), JSON.stringify(pkg, null, 2) + '\n');
    const json = execFileSync('npm', ['pack', '--json', '--pack-destination', outDir], {
      cwd: staging,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'inherit'],
    });
    const [info] = JSON.parse(json);
    return {
      name: pkg.name,
      version: pkg.version,
      filename: info.filename,
      files: info.files.map((f) => f.path),
    };
  } finally {
    rmSync(staging, { recursive: true, force: true });
  }
}

const core = stageAndPack(join(root, 'packages/core'), { rename: CORE_PUBLISH_NAME });
const react = stageAndPack(join(root, 'packages/react'), { requireFiles: ['dist/styles.css'] });

for (const p of [core, react]) {
  console.log(`${p.name}@${p.version} → ${join(outDir, p.filename)} (${p.files.length} files)`);
}
if (!react.files.includes('dist/styles.css')) {
  throw new Error('react tarball lacks dist/styles.css');
}
console.log('\nPublish order (orchestrator, OTP):');
console.log(
  `  npm publish ${join(outDir, core.filename)} --access public --tag pelnora --otp=<code>`
);
console.log(
  `  npm publish ${join(outDir, react.filename)} --access public --tag pelnora --otp=<code>`
);
