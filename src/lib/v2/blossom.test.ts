import { describe, expect, it } from "vitest";
import { verifyEvent, type Event } from "nostr-tools";
import { MAX_REPLICAS, downloadTabBody, prepareTab, sanitizeServers, uploadFrames } from "./blossom";
import { BlobUnavailableError } from "./root";
import { fakeBlossom, localSigner } from "./testUtils";

const A = "https://a.example";
const B = "https://b.example";
const C = "https://c.example";

function decodeAuth(header: string): Event {
  expect(header.startsWith("Nostr ")).toBe(true);
  return JSON.parse(atob(header.slice(6))) as Event;
}

describe("sanitizeServers", () => {
  it("keeps https, allows http only for localhost, drops junk, and dedupes", () => {
    expect(
      sanitizeServers([
        "https://a.example/",
        "https://a.example",
        "http://b.example",
        "http://localhost:3000",
        "javascript:alert(1)",
        "not a url",
        "ftp://c.example",
      ]),
    ).toEqual(["https://a.example", "http://localhost:3000"]);
  });
});

describe("Blossom upload and download", () => {
  it("round-trips a tab body, and the server only ever holds ciphertext", async () => {
    const net = fakeBlossom();
    const body = { cells: { "0:0": { value: "super secret salary" } }, rowCount: 1, colCount: 1 };
    const { frameBytes } = await prepareTab(body);
    const [uploaded] = await uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net.fetch });

    const stored = net.stores.get(A)!.get(uploaded.ref.x)!;
    expect(new TextDecoder("latin1").decode(stored)).not.toContain("secret");
    expect(uploaded.ref.n).toBe(stored.byteLength);
    expect(await downloadTabBody(uploaded.ref, [A], { fetch: net.fetch })).toEqual(body);
  });

  it("signs one authorization that covers every blob in the save", async () => {
    const net = fakeBlossom();
    const signer = localSigner();
    const frames = [
      (await prepareTab({ a: 1 })).frameBytes,
      (await prepareTab({ b: 2 })).frameBytes,
      (await prepareTab({ c: 3 })).frameBytes,
    ];
    const uploaded = await uploadFrames(frames, { servers: [A], signer, fetch: net.fetch });

    const headers = new Set(net.uploads.map((upload) => upload.authorization));
    expect(headers.size).toBe(1);
    const auth = decodeAuth([...headers][0]!);
    expect(verifyEvent(auth)).toBe(true);
    expect(auth.kind).toBe(24242);
    expect(auth.pubkey).toBe(signer.pubkey);
    expect(auth.tags).toContainEqual(["t", "upload"]);
    expect(auth.tags.filter((tag) => tag[0] === "x").map((tag) => tag[1]).sort()).toEqual(
      uploaded.map((blob) => blob.ref.x).sort(),
    );
    const expiry = Number(auth.tags.find((tag) => tag[0] === "expiration")![1]);
    expect(expiry).toBeGreaterThan(Math.floor(Date.now() / 1000));
  });

  it(`copies each blob to at most ${MAX_REPLICAS} servers and skips one that is down`, async () => {
    const net = fakeBlossom();
    net.down.add(A);
    const { frameBytes } = await prepareTab({ x: 1 });
    const [uploaded] = await uploadFrames([frameBytes], { servers: [A, B, C], signer: localSigner(), fetch: net.fetch });

    expect(uploaded.servers).toEqual([B, C]);
    expect(net.stores.get(B)!.has(uploaded.ref.x)).toBe(true);
    expect(net.stores.get(C)!.has(uploaded.ref.x)).toBe(true);
  });

  it("fails clearly when no server accepts the upload", async () => {
    const net = fakeBlossom();
    net.down.add(A);
    const { frameBytes } = await prepareTab({ x: 1 });
    await expect(uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net.fetch })).rejects.toThrow(
      BlobUnavailableError,
    );
  });

  it("refuses to upload with no server configured", async () => {
    const { frameBytes } = await prepareTab({ x: 1 });
    await expect(uploadFrames([frameBytes], { servers: [], signer: localSigner(), fetch: fakeBlossom().fetch })).rejects.toThrow(
      "No Blossom server",
    );
  });

  it("downloads from the next server when the first does not have the blob", async () => {
    const net = fakeBlossom();
    const { frameBytes } = await prepareTab({ hello: "world" });
    const [uploaded] = await uploadFrames([frameBytes], { servers: [B], signer: localSigner(), fetch: net.fetch });
    expect(await downloadTabBody(uploaded.ref, [A, B], { fetch: net.fetch })).toEqual({ hello: "world" });
  });

  it("rejects a blob whose bytes were tampered with", async () => {
    const net = fakeBlossom();
    const { frameBytes } = await prepareTab({ hello: "world" });
    const [uploaded] = await uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net.fetch });
    const stored = net.stores.get(A)!.get(uploaded.ref.x)!;
    stored[10] ^= 0xff;
    await expect(downloadTabBody(uploaded.ref, [A], { fetch: net.fetch })).rejects.toThrow(BlobUnavailableError);
  });

  it("does not contact servers that are not https", async () => {
    const net = fakeBlossom();
    const { frameBytes } = await prepareTab({ x: 1 });
    const [uploaded] = await uploadFrames([frameBytes], { servers: [A], signer: localSigner(), fetch: net.fetch });
    await expect(downloadTabBody(uploaded.ref, ["http://evil.example"], { fetch: net.fetch })).rejects.toThrow(
      "No Blossom server",
    );
  });
});
