import { exec } from 'child_process';
import { writeFile, unlink } from 'fs/promises';
import { tmpdir } from 'os';
import { join } from 'path';

/**
 * Write an AppleScript to a temp file and execute it with osascript.
 * Using a file (rather than -e) allows complex multi-line scripts.
 */
export async function runAppleScript(script: string): Promise<string> {
  const id = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const tmpFile = join(tmpdir(), `outlook-mcp-${id}.applescript`);
  try {
    await writeFile(tmpFile, script, 'utf8');
    return await new Promise<string>((resolve, reject) => {
      exec(`osascript "${tmpFile}"`, { maxBuffer: 10 * 1024 * 1024 }, (err, stdout, stderr) => {
        if (err) {
          // osascript errors go to stderr; include both for debugging
          reject(new Error(stderr.trim() || err.message));
        } else {
          resolve(stdout.trim());
        }
      });
    });
  } finally {
    await unlink(tmpFile).catch(() => {});
  }
}

/**
 * Escape a plain string value for safe embedding inside an AppleScript string literal.
 * Returns an AppleScript expression that evaluates to the string.
 * Handles double quotes and newlines.
 */
export function asString(s: string): string {
  // Fast path: no special chars
  if (!/["'\n\r\t]/.test(s)) return `"${s}"`;

  // Build up a concatenation of safe segments and character id escapes
  const parts: string[] = [];
  let seg = '';
  for (const ch of s) {
    if (ch === '"') {
      if (seg) {
        parts.push(`"${seg}"`);
        seg = '';
      }
      parts.push('(character id 34)');
    } else if (ch === '\n') {
      if (seg) {
        parts.push(`"${seg}"`);
        seg = '';
      }
      parts.push('(character id 10)');
    } else if (ch === '\r') {
      if (seg) {
        parts.push(`"${seg}"`);
        seg = '';
      }
      parts.push('(character id 13)');
    } else if (ch === '\t') {
      if (seg) {
        parts.push(`"${seg}"`);
        seg = '';
      }
      parts.push('(character id 9)');
    } else {
      seg += ch;
    }
  }
  if (seg) parts.push(`"${seg}"`);
  if (parts.length === 0) return '""';
  return parts.join(' & ');
}

/**
 * Parse TSV output from AppleScript (tab-separated fields, one record per line).
 * Blank lines are skipped.
 */
export function parseTSV(output: string): string[][] {
  if (!output) return [];
  return output
    .split('\n')
    .filter((line) => line.trim().length > 0)
    .map((line) => line.split('\t'));
}

/**
 * Strip tabs and newlines from a string (for embedding in TSV fields).
 */
export function flattenForTSV(s: string): string {
  return s.replace(/[\t\n\r]/g, ' ');
}
