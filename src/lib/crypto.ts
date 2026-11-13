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
import { ROOT_VERSION, buildRootV2, isSheetRootV2, parseRootV2, serializeRoot } from "./v2/root";

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

/**
 * Builds the sheet event. New events are v2 roots (compressed, inline storage),
 * so sheets far larger than the 64 KB NIP-44 limit as plain JSON can be saved.
 * Throws SheetTooLargeError when even the compressed root does not fit.
 */
export async function buildSheetEvent(params: {
  sheetId: string;
  document: SheetDocument;
  viewKey: string;
  editKey: string;
  createdAt?: number;
}): Promise<Event> {
  const createdAt = params.createdAt ?? Math.floor(Date.now() / 1000);
  const root = serializeRoot(await buildRootV2(params.document));
  const template: EventTemplate = {
    kind: KIND_SHEET,
    created_at: createdAt,
    tags: [
      ["d", params.sheetId],
      ["t", "formstr-sheet"],
      ["client", CLIENT_TAG],
      ["encrypted", "nip44"],
      ["v", String(ROOT_VERSION)],
      ["storage", "inline"],
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
): Promise<SheetDocument> {
  const parsed: unknown = JSON.parse(decryptWithSecret(event.content, viewKey));
  if (isSheetRootV2(parsed)) {
    return normalizeWorkbook(await parseRootV2(parsed), fallbackTitle);
  }
  return normalizeWorkbook(parsed, fallbackTitle);
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
