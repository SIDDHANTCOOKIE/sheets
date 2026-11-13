import {
  finalizeEvent,
  generateSecretKey,
  getPublicKey,
  nip44,
  type Event,
  type EventTemplate,
} from "nostr-tools";
import { bytesToHex, hexToBytes } from "nostr-tools/utils";
import { KIND_SHEET, CLIENT_TAG } from "./constants";
import type { SheetDocument } from "./types";
import { normalizeWorkbook } from "./workbook";
import type { FileSigner } from "@formstr/drive-sdk";
import {
  INLINE_MAX_BYTES,
  ROOT_VERSION,
  SheetTooLargeError,
  buildBlossomRoot,
  buildInlineRoot,
  isSheetRootV2,
  parseRootV2,
  rootSize,
  serializeRoot,
  splitTab,
  storageOf,
  type BlobRef,
  type SheetRootV2,
} from "./v2/root";
import { downloadTabBody, prepareTab, sanitizeServers, uploadFrames, type BlossomEnv } from "./v2/blossom";

/** A blob already stored on Blossom, remembered so an unchanged tab is not uploaded again. */
export interface ReuseEntry {
  ref: BlobRef;
  servers: string[];
}

export interface BlossomContext {
  /** The user's Blossom servers (their kind 10063 list). */
  servers: string[];
  /** Signs the upload authorization. Required to save sheets that do not fit inline. */
  signer?: FileSigner;
  /**
   * Blobs known to be on Blossom for this sheet, by tab id. Read to skip unchanged
   * tabs, and updated after a successful save or load.
   */
  reuse?: Map<string, ReuseEntry>;
  env?: BlossomEnv;
}

export interface DocKeys {
  viewKey: string;
  editKey: string;
  editPubkey: string;
}

export function generateDocKeys(): DocKeys {
  const view = generateSecretKey();
  const edit = generateSecretKey();
  return {
    viewKey: bytesToHex(view),
    editKey: bytesToHex(edit),
    editPubkey: getPublicKey(edit),
  };
}

export function conversationKeyForSecret(secretHex: string): Uint8Array {
  const secret = hexToBytes(secretHex);
  return nip44.v2.utils.getConversationKey(secret, getPublicKey(secret));
}

export function encryptWithSecret(plaintext: string, secretHex: string): string {
  return nip44.v2.encrypt(plaintext, conversationKeyForSecret(secretHex));
}

export function decryptWithSecret(ciphertext: string, secretHex: string): string {
  return nip44.v2.decrypt(ciphertext, conversationKeyForSecret(secretHex));
}

/** Uploads whichever tabs changed to Blossom and returns a root that points at them. */
async function buildBlossomRootFor(
  doc: SheetDocument,
  context: BlossomContext | undefined,
  inlineBytes: number,
): Promise<SheetRootV2> {
  const servers = sanitizeServers(context?.servers ?? []);
  if (!context?.signer || servers.length === 0) throw new SheetTooLargeError(inlineBytes);

  const prepared = await Promise.all(doc.sheets.map((sheet) => prepareTab(splitTab(sheet).body)));
  const known = context.reuse;
  const entries: (ReuseEntry | undefined)[] = doc.sheets.map((sheet, index) => {
    const entry = known?.get(sheet.id);
    return entry && entry.ref.h === prepared[index].hash ? entry : undefined;
  });

  const missing = entries.flatMap((entry, index) => (entry ? [] : [index]));
  const uploaded = await uploadFrames(
    missing.map((index) => prepared[index].frameBytes),
    { servers, signer: context.signer, ...context.env },
  );
  missing.forEach((index, i) => {
    entries[index] = uploaded[i];
  });

  const complete = entries as ReuseEntry[];
  doc.sheets.forEach((sheet, index) => known?.set(sheet.id, complete[index]));
  const holders = sanitizeServers(complete.flatMap((entry) => entry.servers));
  return buildBlossomRoot(doc, complete.map((entry) => entry.ref), holders);
}

/**
 * Builds the sheet event as a v2 root. A sheet whose compressed root fits in 32 KB
 * stays inline in the event; a larger one has its tabs encrypted and uploaded to
 * the user's Blossom servers, and the event only points at them. Throws
 * SheetTooLargeError when a sheet needs Blossom but no server/signer is available.
 */
export async function buildSheetEvent(params: {
  sheetId: string;
  document: SheetDocument;
  viewKey: string;
  editKey: string;
  createdAt?: number;
  blossom?: BlossomContext;
}): Promise<Event> {
  const createdAt = params.createdAt ?? Math.floor(Date.now() / 1000);
  const inlineRoot = await buildInlineRoot(params.document);
  const inlineBytes = rootSize(inlineRoot);
  const rootValue =
    inlineBytes <= INLINE_MAX_BYTES
      ? inlineRoot
      : await buildBlossomRootFor(params.document, params.blossom, inlineBytes);
  const root = serializeRoot(rootValue);
  const template: EventTemplate = {
    kind: KIND_SHEET,
    created_at: createdAt,
    tags: [
      ["d", params.sheetId],
      ["t", "formstr-sheet"],
      ["client", CLIENT_TAG],
      ["encrypted", "nip44"],
      ["v", String(ROOT_VERSION)],
      ["storage", storageOf(rootValue)],
    ],
    content: encryptWithSecret(root, params.viewKey),
  };
  return finalizeEvent(template, hexToBytes(params.editKey));
}

/**
 * Reads a sheet event of any version. v2 roots are expanded; anything else is
 * treated as a v1 event, whose content is the workbook JSON itself (including
 * the older fortune-sheet format that normalizeWorkbook still understands).
 */
export async function decodeSheetEvent(
  event: Pick<Event, "content">,
  viewKey: string,
  fallbackTitle?: string,
  blossom?: Pick<BlossomContext, "servers" | "reuse" | "env">,
): Promise<SheetDocument> {
  const parsed: unknown = JSON.parse(decryptWithSecret(event.content, viewKey));
  if (!isSheetRootV2(parsed)) return normalizeWorkbook(parsed, fallbackTitle);

  // The root names its own servers; the reader's own list is a fallback.
  const doc = await parseRootV2(parsed, (ref, rootServers) =>
    downloadTabBody(ref, [...rootServers, ...(blossom?.servers ?? [])], blossom?.env),
  );
  for (const tab of parsed.tabs) {
    if (tab.blob) blossom?.reuse?.set(tab.id, { ref: tab.blob, servers: parsed.servers });
  }
  return normalizeWorkbook(doc, fallbackTitle);
}

export function buildSheetDeletionEvent(params: {
  sheetId: string;
  editKey: string;
  eventId?: string;
  createdAt?: number;
}): Event {
  const tags: string[][] = [["a", `${KIND_SHEET}:${getPublicKey(hexToBytes(params.editKey))}:${params.sheetId}`]];
  if (params.eventId) tags.push(["e", params.eventId]);
  const template: EventTemplate = {
    kind: 5,
    created_at: params.createdAt ?? Math.floor(Date.now() / 1000),
    tags,
    content: "Delete spreadsheet",
  };
  return finalizeEvent(template, hexToBytes(params.editKey));
}
