// Encrypted Blossom storage for worksheet payloads.
//
// Encryption and HTTP are @formstr/drive-sdk's (the NIP-FS segmented AES-GCM
// format Formstr Drive uses), so the blob format is shared across Formstr apps.

import {
  createBlossomAuthorization,
  createFetchBlossomTransport,
  decryptFileBytes,
  encryptFile,
  type FileSigner,
} from "@formstr/drive-sdk";
import { frame, unframe } from "./codec";
import { BlobUnavailableError, type BlobRef } from "./root";

/**
 * A blob is copied to this many of the user's servers unless they opt in to more. Every extra
 * copy repeats the whole upload, which slows saves of big sheets, so one is the default.
 */
export const MAX_REPLICAS = 1;
/** Copies when the user turns on "keep a second copy" (and has two servers ticked). */
export const REPLICAS_WHEN_ON = 2;
/** Blossom upload authorizations are short-lived: one is signed per save. */
const AUTH_TTL_SECONDS = 600;

export interface BlossomEnv {
  fetch?: typeof fetch;
  now?: () => number;
}

export interface UploadedBlob {
  ref: BlobRef;
  /** Servers that accepted this blob. */
  servers: string[];
}

/** Only https servers (or http for localhost development) are ever contacted. */
export function sanitizeServers(servers: readonly string[]): string[] {
  const out: string[] = [];
  for (const raw of servers) {
    try {
      const url = new URL(raw);
      const local = url.hostname === "localhost" || url.hostname === "127.0.0.1";
      if (url.protocol === "https:" || (url.protocol === "http:" && local)) {
        const clean = url.origin + url.pathname.replace(/\/+$/, "");
        if (!out.includes(clean)) out.push(clean);
      }
    } catch {
      // not a URL
    }
  }
  return out;
}

/**
 * The list to publish: what is already on the profile, minus what the user removed,
 * plus what they added. Starting from the current list means a server added in another
 * app, or one a slow relay failed to return, is not silently dropped by an edit.
 */
export function mergeServers(current: readonly string[], added: readonly string[], removed: readonly string[]): string[] {
  const dropped = new Set(sanitizeServers(removed));
  return sanitizeServers([...sanitizeServers(current).filter((url) => !dropped.has(url)), ...added]);
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", bytes as BufferSource);
  return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, "0")).join("");
}

/** The plaintext frame for a tab body, with its hash, ready to compare or upload. */
export async function prepareTab(body: unknown): Promise<{ frameBytes: Uint8Array; hash: string }> {
  const frameBytes = await frame(body);
  return { frameBytes, hash: await sha256Hex(frameBytes) };
}

/**
 * Encrypts and uploads frames. One authorization covers every blob (BUD-11
 * multi-x), so a save costs a single signer prompt however many tabs changed.
 * Each blob goes to up to MAX_REPLICAS servers and must land on at least one.
 */
export async function uploadFrames(
  frames: Uint8Array[],
  options: { servers: readonly string[]; signer: FileSigner; replicas?: number } & BlossomEnv,
): Promise<UploadedBlob[]> {
  const servers = sanitizeServers(options.servers);
  if (servers.length === 0) throw new BlobUnavailableError("No Blossom server is configured.");
  if (frames.length === 0) return [];

  const copies = Math.max(1, options.replicas ?? MAX_REPLICAS);
  const now = options.now ?? (() => Math.floor(Date.now() / 1000));
  const transport = createFetchBlossomTransport(options.fetch ?? fetch);
  const encrypted = await Promise.all(frames.map((bytes) => encryptFile(bytes)));
  const authorization = await createBlossomAuthorization(
    options.signer,
    "upload",
    encrypted.map((blob) => blob.blobHash),
    "Upload spreadsheet data",
    AUTH_TTL_SECONDS,
    now,
  );

  const uploaded: UploadedBlob[] = [];
  for (const blob of encrypted) {
    const accepted: string[] = [];
    const failures: string[] = [];
    for (const server of servers) {
      if (accepted.length >= copies) break;
      try {
        await transport.upload({ server, bytes: blob.bytes, authorization });
        accepted.push(server);
      } catch (error) {
        failures.push(`${server}: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
    if (accepted.length === 0) {
      throw new BlobUnavailableError(`Couldn't upload to any Blossom server. ${failures.join("; ")}`);
    }
    uploaded.push({
      servers: accepted,
      ref: {
        x: blob.blobHash,
        k: blob.encryptionKey,
        n: blob.bytes.byteLength,
        p: blob.size,
        c: blob.chunkSize,
        h: blob.unencryptedFileHash,
      },
    });
  }
  return uploaded;
}

/** Downloads, verifies and decrypts one blob, trying each server until one has it. */
export async function downloadTabBody(
  ref: BlobRef,
  servers: readonly string[],
  env: BlossomEnv = {},
): Promise<unknown> {
  const candidates = sanitizeServers(servers);
  if (candidates.length === 0) throw new BlobUnavailableError("No Blossom server is available to load this sheet from.");
  const transport = createFetchBlossomTransport(env.fetch ?? fetch);
  const failures: string[] = [];
  for (const server of candidates) {
    try {
      const encrypted = await transport.download({ server, hash: ref.x, expectedSize: ref.n });
      const plaintext = await decryptFileBytes(encrypted, {
        size: ref.p,
        chunkSize: ref.c,
        blobHash: ref.x,
        encryptionKey: ref.k,
        unencryptedFileHash: ref.h,
        servers: candidates,
        type: "application/octet-stream",
      });
      return await unframe(plaintext);
    } catch (error) {
      failures.push(`${server}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  throw new BlobUnavailableError(`Couldn't load this sheet from any Blossom server. ${failures.join("; ")}`);
}
