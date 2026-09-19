/* eslint-disable @typescript-eslint/no-var-requires */
declare const __dirname: string;

const fs = require('fs') as {
  readFileSync(fileName: string, encoding: string): string;
  readdirSync(directory: string, options: { withFileTypes: boolean }): Array<{ name: string; isDirectory(): boolean }>;
};
const path = require('path') as {
  join(...parts: string[]): string;
  resolve(...parts: string[]): string;
};
const sourceRoot = path.resolve(__dirname, '../../../../src/webparts/myCalendars');

function captureGroups(source: string, expression: RegExp): string[] {
  const values: string[] = [];
  let match: RegExpExecArray | null;
  while ((match = expression.exec(source)) !== null) {
    values.push(match[1]);
  }
  return values.sort();
}

function readKeys(fileName: string, expression: RegExp): string[] {
  const source = fs.readFileSync(path.join(sourceRoot, 'loc', fileName), 'utf8');
  return captureGroups(source, expression);
}

function getSourceFiles(directory: string): string[] {
  const files: string[] = [];
  fs.readdirSync(directory, { withFileTypes: true }).forEach(entry => {
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      files.push(...getSourceFiles(fullPath));
    } else if (/\.(ts|tsx)$/.test(entry.name) && !/\.test\.(ts|tsx)$/.test(entry.name)) {
      files.push(fullPath);
    }
  });
  return files;
}

describe('localization resource consistency', () => {
  it('keeps the declaration and locale resource keys synchronized', () => {
    const declarationKeys = readKeys('mystrings.d.ts', /^\s{2}([A-Za-z0-9_]+): string;/gm);
    const englishKeys = readKeys('en-us.js', /["']([A-Za-z0-9_]+)["']\s*:/g);
    const dutchKeys = readKeys('nl-nl.js', /["']([A-Za-z0-9_]+)["']\s*:/g);

    expect(englishKeys).toEqual(declarationKeys);
    expect(dutchKeys).toEqual(declarationKeys);
  });

  it('declares every statically referenced strings key in both locales', () => {
    const source = getSourceFiles(sourceRoot)
      .map(fileName => fs.readFileSync(fileName, 'utf8'))
      .join('\n');
    const usedKeys = captureGroups(source, /strings\.([A-Za-z0-9_]+)/g);
    const declarationKeys = new Set(readKeys('mystrings.d.ts', /^\s{2}([A-Za-z0-9_]+): string;/gm));
    const englishKeys = new Set(readKeys('en-us.js', /["']([A-Za-z0-9_]+)["']\s*:/g));
    const dutchKeys = new Set(readKeys('nl-nl.js', /["']([A-Za-z0-9_]+)["']\s*:/g));

    expect(usedKeys.filter(key => !declarationKeys.has(key))).toEqual([]);
    expect(usedKeys.filter(key => !englishKeys.has(key))).toEqual([]);
    expect(usedKeys.filter(key => !dutchKeys.has(key))).toEqual([]);
  });
});
