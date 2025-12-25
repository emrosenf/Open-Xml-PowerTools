/**
 * Hashing utilities for content comparison
 *
 * Uses js-sha256 for pure JavaScript SHA-256 implementation.
 * Works in browser, Node.js, and WASM environments.
 */

import { sha256 } from 'js-sha256';

/**
 * Compute SHA-256 hash of a string
 */
export function hashString(str: string): string {
  return sha256(str);
}

/**
 * Compute SHA-256 hash of binary data
 */
export function hashBytes(data: Uint8Array | ArrayBuffer): string {
  return sha256(data);
}

/**
 * Compute a short hash (first 8 characters) for display/debugging
 */
export function shortHash(str: string): string {
  return sha256(str).substring(0, 8);
}

/**
 * Generate a unique ID from content
 */
export function contentId(content: string): string {
  return sha256(content).substring(0, 16);
}

/**
 * Compare two hashes for equality
 */
export function hashEquals(hash1: string, hash2: string): boolean {
  return hash1 === hash2;
}
