import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { mkdirSync, writeFileSync, rmSync, chmodSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { discoverFiles } from './fileDiscovery.js';

const TEST_DIR = join(import.meta.dirname, '__test_fixtures_discovery');

function createFile(relativePath: string, content: string | Buffer): void {
  const fullPath = join(TEST_DIR, relativePath);
  mkdirSync(dirname(fullPath), { recursive: true });
  writeFileSync(fullPath, content);
}

beforeEach(() => {
  mkdirSync(TEST_DIR, { recursive: true });
});

afterEach(() => {
  rmSync(TEST_DIR, { recursive: true, force: true });
});

describe('discoverFiles', () => {
  it('collects .bas, .cls, and .frm files recursively', () => {
    createFile('root.bas', 'Attribute VB_Name = "RootMod"\nPublic Sub Main()\nEnd Sub');
    createFile('sub/nested.cls', 'Attribute VB_Name = "NestedClass"\nOption Explicit');
    createFile('sub/deep/form.frm', 'Attribute VB_Name = "MyForm"\nBegin VB.Form');

    const files = discoverFiles(TEST_DIR);
    expect(files).toHaveLength(3);

    const types = files.map(f => f.type).sort();
    expect(types).toEqual(['bas', 'cls', 'frm']);
  });

  it('ignores non-VB6 files', () => {
    createFile('code.bas', 'Attribute VB_Name = "Code"\nPublic Sub X()\nEnd Sub');
    createFile('readme.txt', 'This is a readme');
    createFile('data.json', '{}');

    const files = discoverFiles(TEST_DIR);
    expect(files).toHaveLength(1);
    expect(files[0].type).toBe('bas');
  });

  it('extracts moduleName from Attribute VB_Name line', () => {
    createFile('TCP.bas', 'Attribute VB_Name = "Mod_TCP"\nOption Explicit');

    const files = discoverFiles(TEST_DIR);
    expect(files).toHaveLength(1);
    expect(files[0].moduleName).toBe('Mod_TCP');
  });

  it('falls back to filename without extension when Attribute VB_Name is missing', () => {
    createFile('MyModule.bas', 'Option Explicit\nPublic Sub Main()\nEnd Sub');

    const files = discoverFiles(TEST_DIR);
    expect(files).toHaveLength(1);
    expect(files[0].moduleName).toBe('MyModule');
  });

  it('returns relative paths from rootDir', () => {
    createFile('sub/nested.bas', 'Attribute VB_Name = "Nested"\nDim x As Long');

    const files = discoverFiles(TEST_DIR);
    expect(files).toHaveLength(1);
    expect(files[0].path).toBe(join('sub', 'nested.bas'));
  });

  it('reads file content correctly', () => {
    const content = 'Attribute VB_Name = "Test"\nPublic Sub Hello()\n  MsgBox "Hi"\nEnd Sub';
    createFile('test.bas', content);

    const files = discoverFiles(TEST_DIR);
    expect(files).toHaveLength(1);
    expect(files[0].content).toBe(content);
  });

  it('handles latin1 fallback for Windows-1252 encoded files', () => {
    // Create a file with bytes that are invalid UTF-8 but valid in latin1/Windows-1252
    // 0xF1 = ñ in latin1, which is part of a multi-byte sequence in UTF-8
    const latin1Bytes = Buffer.from([
      0x41, 0x74, 0x74, 0x72, 0x69, 0x62, 0x75, 0x74, 0x65, 0x20, // "Attribute "
      0x56, 0x42, 0x5F, 0x4E, 0x61, 0x6D, 0x65, 0x20, 0x3D, 0x20, // "VB_Name = "
      0x22, 0x54, 0x65, 0x73, 0x74, 0x22, 0x0A,                     // '"Test"\n'
      0x27, 0x20, 0x43, 0x6F, 0x6D, 0x65, 0x6E, 0x74, 0x61, 0x72, // "' Comentar"
      0x69, 0x6F, 0x20, 0x65, 0x6E, 0x20, 0x65, 0x73, 0x70, 0x61, // "io en espa"
      0xF1, 0x6F, 0x6C,                                               // "ñol" (ñ as 0xF1)
    ]);
    createFile('latin1.bas', latin1Bytes);

    const files = discoverFiles(TEST_DIR);
    expect(files).toHaveLength(1);
    expect(files[0].moduleName).toBe('Test');
    // latin1 decoding should preserve the ñ as U+00F1
    expect(files[0].content).toContain('espa\u00F1ol');
  });

  it('skips unreadable files with a warning', () => {
    createFile('good.bas', 'Attribute VB_Name = "Good"\nDim x As Long');

    // Create a directory where we expect a file — this will cause readFileSync to fail
    mkdirSync(join(TEST_DIR, 'unreadable'), { recursive: true });

    const files = discoverFiles(TEST_DIR);
    // Should still find the good file
    expect(files).toHaveLength(1);
    expect(files[0].moduleName).toBe('Good');
  });

  it('returns empty array for empty directory', () => {
    const files = discoverFiles(TEST_DIR);
    expect(files).toHaveLength(0);
  });

  it('handles Attribute VB_Name with extra spaces', () => {
    createFile('spaced.bas', 'Attribute  VB_Name  =  "SpacedMod"\nOption Explicit');

    const files = discoverFiles(TEST_DIR);
    expect(files).toHaveLength(1);
    // The regex uses \s+ so extra spaces should still match
    expect(files[0].moduleName).toBe('SpacedMod');
  });
});
