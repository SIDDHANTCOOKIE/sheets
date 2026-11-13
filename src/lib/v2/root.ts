// Sheets v2 root manifest (kind 33460), inline storage mode.
//
// Each tab's cells and layout are stored as one compressed FSB2 payload inside
// the root, so a sheet that is large as JSON still fits in a single event.
// Larger sheets will move their payloads to Blossom blobs in a later change;
// the root shape already leaves room for that (`servers`, per-tab payloads).

import { WORKBOOK_SCHEMA, type WorkbookDocument, type WorkbookSheet } from "../workbook";
import { base64ToBytes, bytesToBase64, frame, unframe } from "./codec";

export const ROOT_SCHEMA = "formstr.sheets.root";
export const ROOT_VERSION = 2;

/**
 * Largest root plaintext we will encrypt. The binding limit is the relay's
 * event size, not the cipher: relay.formstr.app caps event content at 102,400
 * bytes and nos.lol caps messages at 131,072 (both from their NIP-11 documents).
 * NIP-44 ciphertext is base64, so ~4/3 of the plaintext. 48 KB encrypts to about
 * 65 KB, which fits those relays and also stays inside the 65,535-byte plaintext
 * limit of the standard NIP-44 spec, so other NIP-44 clients can still decrypt it
 * (nostr-tools' extended format for larger payloads is not interoperable).
 */
export const INLINE_MAX_BYTES = 48 * 1024;

type TabMeta = Pick<WorkbookSheet, "id" | "name" | "order" | "hidden" | "color" | "showGridLines">;

export interface RootTab extends TabMeta {
  /** base64 of an FSB2 payload holding the rest of the worksheet. */
  inline: string;
}

export interface SheetRootV2 {
  schema: typeof ROOT_SCHEMA;
  v: typeof ROOT_VERSION;
  title: string;
  activeSheetId: string;
  createdAt: number;
  updatedAt: number;
  epoch: number;
  servers: string[];
  tabs: RootTab[];
}

export class SheetTooLargeError extends Error {
  readonly bytes: number;
  readonly limit: number;
  constructor(bytes: number, limit = INLINE_MAX_BYTES) {
    super(
      `This spreadsheet is too large to save yet (${Math.ceil(bytes / 1024)} KB compressed, limit ${Math.floor(limit / 1024)} KB). Larger sheets need Blossom storage, which is coming soon.`,
    );
    this.name = "SheetTooLargeError";
    this.bytes = bytes;
    this.limit = limit;
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

export async function buildRootV2(doc: WorkbookDocument): Promise<SheetRootV2> {
  const tabs: RootTab[] = [];
  for (const sheet of doc.sheets) {
    const { id, name, order, hidden, color, showGridLines, ...body } = sheet;
    const payload = await frame(body);
    tabs.push({ id, name, order, hidden, color, showGridLines, inline: bytesToBase64(payload) });
  }
  return {
    schema: ROOT_SCHEMA,
    v: ROOT_VERSION,
    title: doc.title,
    activeSheetId: doc.activeSheetId,
    createdAt: doc.createdAt,
    updatedAt: doc.updatedAt,
    epoch: 1,
    servers: [],
    tabs,
  };
}

/** Serialises a root, refusing roots that would not fit in one encrypted event. */
export function serializeRoot(root: SheetRootV2): string {
  const json = JSON.stringify(root);
  const bytes = new TextEncoder().encode(json).length;
  if (bytes > INLINE_MAX_BYTES) throw new SheetTooLargeError(bytes);
  return json;
}

export async function parseRootV2(root: SheetRootV2): Promise<WorkbookDocument> {
  const sheets: WorkbookSheet[] = [];
  for (const tab of root.tabs) {
    const body = (await unframe(base64ToBytes(tab.inline))) as Omit<WorkbookSheet, keyof TabMeta>;
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
