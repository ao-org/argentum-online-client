#!/usr/bin/env node

/**
 * Language File Validator
 * 
 * Validates that all language JSON files (1.json through 5.json) have
 * the exact same set of keys. Uses 1.json (Spanish) as the source of truth.
 * 
 * Exit code 0 = all good, exit code 1 = mismatches found.
 */

const fs = require('fs');
const path = require('path');

const LANGUAGES_DIR = path.join(__dirname, '..', 'Languages');
const SOURCE_FILE = '1.json'; // Spanish = source of truth
const LANG_FILES = ['1.json', '2.json', '3.json', '4.json', '5.json'];
const LANG_NAMES = { '1.json': 'Spanish', '2.json': 'English', '3.json': 'Portuguese', '4.json': 'French', '5.json': 'Italian' };

let hasErrors = false;

// 1. Parse all files
const parsed = {};
for (const file of LANG_FILES) {
    const filePath = path.join(LANGUAGES_DIR, file);
    if (!fs.existsSync(filePath)) {
        console.error(`FAIL: ${file} (${LANG_NAMES[file]}) not found!`);
        hasErrors = true;
        continue;
    }
    try {
        parsed[file] = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    } catch (e) {
        console.error(`FAIL: ${file} (${LANG_NAMES[file]}) has invalid JSON: ${e.message}`);
        hasErrors = true;
    }
}

if (hasErrors) {
    process.exit(1);
}

// 2. Get source keys
const sourceKeys = new Set(Object.keys(parsed[SOURCE_FILE]));
console.log(`Source: ${SOURCE_FILE} (${LANG_NAMES[SOURCE_FILE]}) has ${sourceKeys.size} keys.\n`);

// 3. Compare each file against source
for (const file of LANG_FILES) {
    if (file === SOURCE_FILE) continue;

    const fileKeys = new Set(Object.keys(parsed[file]));
    const missing = [...sourceKeys].filter(k => !fileKeys.has(k));
    const extra = [...fileKeys].filter(k => !sourceKeys.has(k));

    if (missing.length === 0 && extra.length === 0) {
        console.log(`OK: ${file} (${LANG_NAMES[file]}) - ${fileKeys.size} keys, all match.`);
    } else {
        hasErrors = true;
        console.error(`FAIL: ${file} (${LANG_NAMES[file]})`);
        if (missing.length > 0) {
            console.error(`  Missing ${missing.length} key(s):`);
            missing.forEach(k => console.error(`    - ${k}`));
        }
        if (extra.length > 0) {
            console.error(`  Extra ${extra.length} key(s) not in source:`);
            extra.forEach(k => console.error(`    - ${k}`));
        }
    }
}

// 4. Check for duplicate keys in raw text (JSON.parse silently drops dupes)
console.log('');
for (const file of LANG_FILES) {
    const raw = fs.readFileSync(path.join(LANGUAGES_DIR, file), 'utf8');
    const keyRegex = /"([^"]+)"\s*:/g;
    const seen = {};
    let match;
    while ((match = keyRegex.exec(raw)) !== null) {
        const key = match[1];
        if (seen[key]) {
            console.error(`FAIL: ${file} (${LANG_NAMES[file]}) has duplicate key: "${key}"`);
            hasErrors = true;
        }
        seen[key] = true;
    }
}

if (hasErrors) {
    console.error('\nLanguage validation FAILED.');
    process.exit(1);
} else {
    console.log('\nAll language files are in sync.');
    process.exit(0);
}
