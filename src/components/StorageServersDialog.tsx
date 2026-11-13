import { useEffect, useState } from "react";
import { HardDrive, Loader2, Plus, X } from "lucide-react";
import { Dialog } from "./Dialog";
import { useLibrary } from "../providers/LibraryProvider";
import { useToast } from "../providers/ToastProvider";
import { sanitizeServers } from "../lib/v2/blossom";

/**
 * Where large sheets are stored. This edits the user's Blossom server list (a Nostr
 * kind 10063 event), the same list other Nostr apps such as Formstr Drive use.
 */
export function StorageServersDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  const library = useLibrary();
  const { toast } = useToast();
  const [loaded, setLoaded] = useState<string[] | null>(null);
  const [servers, setServers] = useState<string[]>([]);
  const [draft, setDraft] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    if (!open) return;
    let cancelled = false;
    setLoaded(null);
    setDraft("");
    setError("");
    library
      .refreshBlossomServers()
      .then((list) => {
        if (cancelled) return;
        setLoaded(list);
        setServers(list);
      })
      .catch(() => {
        if (cancelled) return;
        setLoaded([]);
        setServers([]);
      });
    return () => {
      cancelled = true;
    };
    // Reload only when the dialog opens.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  const add = () => {
    const [clean] = sanitizeServers([draft.trim()]);
    if (!clean) {
      setError("Enter a server address that starts with https://");
      return;
    }
    setError("");
    if (!servers.includes(clean)) setServers([...servers, clean]);
    setDraft("");
  };

  const save = async () => {
    // A pending, valid address in the box counts as added.
    const pending = sanitizeServers([draft.trim()])[0];
    const next = pending && !servers.includes(pending) ? [...servers, pending] : servers;
    const removed = (loaded ?? []).filter((url) => !next.includes(url));
    setSaving(true);
    setError("");
    try {
      await library.saveBlossomServers(next, removed);
      toast("Storage servers saved.", "success");
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Couldn't save your servers.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Dialog
      open={open}
      onClose={saving ? () => {} : onClose}
      title="Storage servers"
      subtitle="Large spreadsheets are encrypted and stored on Blossom servers you choose."
      size="md"
      footer={
        <div className="dialog-footer-row">
          <span />
          <div className="dialog-footer-actions">
            <button type="button" className="btn btn-ghost" disabled={saving} onClick={onClose}>
              Cancel
            </button>
            <button type="button" className="btn btn-primary" disabled={saving || loaded === null} onClick={save}>
              {saving ? <Loader2 size={16} className="spin" /> : <HardDrive size={16} />}
              Save servers
            </button>
          </div>
        </div>
      }
    >
      <div className="new-sheet-panel">
        <div className="field">
          <span className="field-label">Your servers</span>
          {loaded === null ? (
            <div className="server-empty">
              <Loader2 size={14} className="spin" /> Looking up your server list…
            </div>
          ) : servers.length === 0 ? (
            <div className="server-empty">No servers yet. Add one below to save large sheets.</div>
          ) : (
            <ul className="server-list">
              {servers.map((url) => (
                <li key={url}>
                  <span>{url}</span>
                  <button
                    type="button"
                    className="icon-btn"
                    aria-label={`Remove ${url}`}
                    onClick={() => setServers(servers.filter((item) => item !== url))}
                  >
                    <X size={15} />
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>

        <label className="field">
          <span className="field-label">Add a server</span>
          <div className="server-add">
            <input
              value={draft}
              placeholder="https://blossom.example.com"
              onChange={(event) => setDraft(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault();
                  add();
                }
              }}
            />
            <button type="button" className="btn btn-outline" onClick={add} disabled={!draft.trim()}>
              <Plus size={15} /> Add
            </button>
          </div>
        </label>

        <p className="server-note">
          This is saved to your Nostr profile and shared with other apps that use Blossom. Saving asks your signer to
          approve one event.
        </p>

        {error && <div className="form-error">{error}</div>}
      </div>
    </Dialog>
  );
}
