// Sheets v2 root manifest (kind 33460).
//
// The root is a small encrypted event. Each worksheet (tab) keeps its cells and
// layout in a compressed FSB2 payload that is stored one of two ways:
//
//   inline  - the payload is embedded in the root (small sheets, no server needed)
//   blossom - the payload is an encrypted Blossom blob; the root holds a BlobRef
//
// This file is pure: it never touches the network. Fetching blobs is passed in.

import { WORKBOOK_SCHEMA, type WorkbookDocument, type WorkbookSheet } from "../workbook";
import { base64ToBytes, bytesToBase64, frame, unframe } from "./codec";

export const ROOT_SCHEMA = "formstr.sheets.root";
export const ROOT_VERSION = 2;

/**
 * Largest root plaintext we will ever encrypt. The binding limit is the relay's
 * event size, not the cipher: relay.formstr.app caps event content at 102,400
 * bytes and nos.lol caps messages at 131,072 (both from their NIP-11 documents).
 * NIP-44 ciphertext is base64, so ~4/3 of the plaintext. 48 KB encrypts to about
 * 65 KB, which fits those relays and also stays inside the 65,535-byte plaintext
 * limit of the standard NIP-44 spec, so other NIP-44 clients can still decrypt it
 * (nostr-tools' extended format for larger payloads is not interoperable).
 */
export const ROOT_MAX_BYTES = 48 * 1024;

/** A root at or under this size stays inline; anything bigger moves to Blossom. */
export const INLINE_MAX_BYTES = 32 * 1024;

/** Pointer to one encrypted Blossom blob. All of it is secret: it sits inside the encrypted root. */
export interface BlobRef {
  /** sha256 of the ciphertext, which is the blob's address on Blossom. */
  x: string;
  /** Per-blob secret, from which the AES key is derived. */
  k: string;
  /** Ciphertext size in bytes. */
  n: number;
  /** Plaintext size in bytes (the FSB2 frame). */
  p: number;
  /** Encryption segment size in bytes. */
  c: number;
  /** sha256 of the plaintext frame; also lets an unchanged tab skip re-upload. */
  h: string;
}

type TabMeta = Pick<WorkbookSheet, "id" | "name" | "order" | "hidden" | "color" | "showGridLines">;
type TabBody = Omit<WorkbookSheet, keyof TabMeta>;

export interface RootTab extends TabMeta {
  /** base64 of an FSB2 payload holding the rest of the worksheet. Inline storage. */
  inline?: string;
  /** Blossom blob holding that payload. Blossom storage. */
  blob?: BlobRef;
}

export interface SheetRootV2 {
  schema: typeof ROOT_SCHEMA;
  v: typeof ROOT_VERSION;
  title: string;
  activeSheetId: string;
  createdAt: number;
  updatedAt: number;
  epoch: number;
  /** Blossom servers that hold this sheet's blobs. Empty for inline sheets. */
  servers: string[];
  tabs: RootTab[];
}

export type SheetStorage = "inline" | "blossom";

export class SheetTooLargeError extends Error {
  readonly bytes: number;
  constructor(bytes: number) {
    super(
      `This spreadsheet is too large to store in a single event (${Math.ceil(bytes / 1024)} KB compressed). ` +
        "Add a Blossom server to your Nostr profile (for example in Formstr Drive) to save larger sheets.",
    );
    this.name = "SheetTooLargeError";
    this.bytes = bytes;
  }
}

export class BlobUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "BlobUnavailableError";
  }
}

export function isSheetRootV2(value: unknown): value is SheetRootV2 {
  return (
    !!value &&
    typeof value === "object" &&
    (value as { schema?: unknown }).schema === ROOT_SCHEMA &&
    (value as { v?: unknown }).v === ROOT_VERSION &&
    Array.isArray((value as { tabs?: unknown }).tabs)
  );
}

export function splitTab(sheet: WorkbookSheet): { meta: TabMeta; body: TabBody } {
  const { id, name, order, hidden, color, showGridLines, ...body } = sheet;
  return { meta: { id, name, order, hidden, color, showGridLines }, body };
}

function rootShell(doc: WorkbookDocument, servers: string[], tabs: RootTab[]): SheetRootV2 {
  return {
    schema: ROOT_SCHEMA,
    v: ROOT_VERSION,
    title: doc.title,
    activeSheetId: doc.activeSheetId,
    createdAt: doc.createdAt,
    updatedAt: doc.updatedAt,
    epoch: 1,
    servers,
    tabs,
  };
}

/** Root with every tab's payload embedded. */
export async function buildInlineRoot(doc: WorkbookDocument): Promise<SheetRootV2> {
  const tabs: RootTab[] = [];
  for (const sheet of doc.sheets) {
    const { meta, body } = splitTab(sheet);
    tabs.push({ ...meta, inline: bytesToBase64(await frame(body)) });
  }
  return rootShell(doc, [], tabs);
}

/** Root whose tabs point at Blossom blobs; `refs` is parallel to `doc.sheets`. */
export function buildBlossomRoot(doc: WorkbookDocument, refs: BlobRef[], servers: string[]): SheetRootV2 {
  const tabs = doc.sheets.map((sheet, index) => ({ ...splitTab(sheet).meta, blob: refs[index] }));
  return rootShell(doc, servers, tabs);
}

export function rootSize(root: SheetRootV2): number {
  return new TextEncoder().encode(JSON.stringify(root)).length;
}

export function storageOf(root: SheetRootV2): SheetStorage {
  return root.tabs.some((tab) => tab.blob) ? "blossom" : "inline";
}

/** Serialises a root, refusing roots that would not fit in one encrypted event. */
export function serializeRoot(root: SheetRootV2): string {
  const json = JSON.stringify(root);
  const bytes = new TextEncoder().encode(json).length;
  if (bytes > ROOT_MAX_BYTES) throw new SheetTooLargeError(bytes);
  return json;
}

export type FetchTabBody = (ref: BlobRef, servers: string[]) => Promise<unknown>;

export async function parseRootV2(root: SheetRootV2, fetchTabBody?: FetchTabBody): Promise<WorkbookDocument> {
  const sheets: WorkbookSheet[] = [];
  for (const tab of root.tabs) {
    let body: TabBody;
    if (tab.inline !== undefined) {
      body = (await unframe(base64ToBytes(tab.inline))) as TabBody;
    } else if (tab.blob) {
      if (!fetchTabBody) throw new BlobUnavailableError("This sheet is stored on Blossom, but no server is available to load it from.");
      body = (await fetchTabBody(tab.blob, root.servers)) as TabBody;
    } else {
      throw new Error(`Worksheet "${tab.name}" has no data.`);
    }
    sheets.push({
      ...body,
      id: tab.id,
      name: tab.name,
      order: tab.order,
      hidden: tab.hidden,
      color: tab.color,
      showGridLines: tab.showGridLines,
    });
  }
  return {
    schema: WORKBOOK_SCHEMA,
    version: 1,
    title: root.title,
    activeSheetId: root.activeSheetId,
    createdAt: root.createdAt,
    updatedAt: root.updatedAt,
    sheets,
  };
}
