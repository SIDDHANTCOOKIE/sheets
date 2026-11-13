import { describe, expect, it } from "vitest";
import { finalizeEvent } from "nostr-tools";
import { hexToBytes } from "nostr-tools/utils";
import {
  buildSheetEvent,
  decodeSheetEvent,
  decryptWithSecret,
  encryptWithSecret,
  generateDocKeys,
} from "./crypto";
import { createEmptyWorkbook } from "./workbook";

describe("sheet crypto", () => {
  it("encrypts and decrypts with a self-derived conversation key", () => {
    const keys = generateDocKeys();
    const ciphertext = encryptWithSecret("private cell contents", keys.viewKey);
    expect(ciphertext).not.toContain("private");
    expect(decryptWithSecret(ciphertext, keys.viewKey)).toBe("private cell contents");
  });

  it("signs sheet events with the edit key and encrypts with the view key", async () => {
    const keys = generateDocKeys();
    const doc = createEmptyWorkbook("Encrypted test");
    const event = await buildSheetEvent({
      sheetId: "test",
      document: doc,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      createdAt: 1,
    });
    expect(event.pubkey).toBe(keys.editPubkey);
    expect(event.kind).toBe(33460);
    expect(event.tags).toContainEqual(["d", "test"]);
    expect(event.tags).toContainEqual(["encrypted", "nip44"]);
    expect(event.tags).toContainEqual(["v", "2"]);
    expect(event.tags).toContainEqual(["storage", "inline"]);
    const decrypted = JSON.parse(decryptWithSecret(event.content, keys.viewKey)) as {
      title: string;
      v: number;
    };
    expect(decrypted.title).toBe("Encrypted test");
    expect(decrypted.v).toBe(2);
    expect((await decodeSheetEvent(event, keys.viewKey)).title).toBe("Encrypted test");
  });

  it("still reads v1 events written before the v2 format", async () => {
    const keys = generateDocKeys();
    const doc = createEmptyWorkbook("Old sheet");
    doc.sheets[0].cells["0:0"] = { value: "legacy", type: "text" };
    // Exactly what the previous buildSheetEvent produced: the workbook JSON as the content.
    const v1Event = finalizeEvent(
      {
        kind: 33460,
        created_at: 1,
        tags: [
          ["d", "old"],
          ["t", "formstr-sheet"],
          ["client", "formstr-sheets"],
          ["encrypted", "nip44"],
        ],
        content: encryptWithSecret(JSON.stringify(doc), keys.viewKey),
      },
      hexToBytes(keys.editKey),
    );
    const decoded = await decodeSheetEvent(v1Event, keys.viewKey);
    expect(decoded.title).toBe("Old sheet");
    expect(decoded.sheets[0].cells["0:0"]).toEqual({ value: "legacy", type: "text" });
  });

  it("still reads v1 events in the older fortune-sheet format", async () => {
    const keys = generateDocKeys();
    const fortuneStyle = {
      title: "Fortune era",
      sheets: [{ id: "s1", name: "Sheet1", status: 1, order: 0, celldata: [{ r: 0, c: 0, v: { v: "hi", m: "hi", ct: { fa: "General", t: "g" } } }] }],
    };
    const event = finalizeEvent(
      {
        kind: 33460,
        created_at: 1,
        tags: [["d", "fortune"]],
        content: encryptWithSecret(JSON.stringify(fortuneStyle), keys.viewKey),
      },
      hexToBytes(keys.editKey),
    );
    const decoded = await decodeSheetEvent(event, keys.viewKey);
    expect(decoded.title).toBe("Fortune era");
    expect(decoded.sheets[0].cells["0:0"]?.value).toBe("hi");
  });

  it("fits a sheet into the relay content limit that the v1 format exceeded", async () => {
    // relay.formstr.app advertises max_content_length 102400 in its NIP-11 document.
    const RELAY_CONTENT_LIMIT = 102_400;
    const keys = generateDocKeys();
    const doc = createEmptyWorkbook("Large");
    const cells: Record<string, { value: string; type: "text" }> = {};
    for (let r = 0; r < 400; r++) {
      for (let c = 0; c < 10; c++) cells[`${r}:${c}`] = { value: `row ${r} col ${c} some repeated text`, type: "text" };
    }
    doc.sheets[0] = { ...doc.sheets[0], rowCount: 400, colCount: 10, cells };

    // What the v1 format published: the workbook JSON encrypted directly.
    const v1Content = encryptWithSecret(JSON.stringify(doc), keys.viewKey);
    expect(v1Content.length).toBeGreaterThan(RELAY_CONTENT_LIMIT);

    const event = await buildSheetEvent({ sheetId: "large", document: doc, viewKey: keys.viewKey, editKey: keys.editKey });
    expect(event.content.length).toBeLessThan(RELAY_CONTENT_LIMIT);
    const decoded = await decodeSheetEvent(event, keys.viewKey);
    expect(Object.keys(decoded.sheets[0].cells)).toHaveLength(4000);
  });
});
