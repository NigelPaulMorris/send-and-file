/* Send & File - shared code used by the task pane and the event handlers. */

const CONFIG = {
  // 1) Paste the "Application (client) ID" from your Entra app registration here.
  clientId: "YOUR-CLIENT-ID",
  // 2) Must match <Id> in manifest.xml (lowercase). Leave as-is unless you change the manifest Id.
  addinId: "fafb1f50-165d-4d79-831f-aa09f013b560",
};

const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["Mail.ReadWrite"];

// Office stores an add-in's custom properties on the email in this named property.
// It survives sending, so the copy in Sent Items still carries the chosen folder.
const CECP_PROP_ID =
  "String {00020329-0000-0000-C000-000000000046} Name cecp-" + CONFIG.addinId;

let _pca = null;
async function getPca() {
  if (!_pca) {
    _pca = await msal.createNestablePublicClientApplication({
      auth: { clientId: CONFIG.clientId, authority: "https://login.microsoftonline.com/common" },
    });
  }
  return _pca;
}

/** Gets a Graph token using Outlook's own sign-in. interactive=true may show a consent popup. */
async function getToken(interactive = false) {
  const pca = await getPca();
  try {
    return (await pca.acquireTokenSilent({ scopes: SCOPES })).accessToken;
  } catch (err) {
    if (!interactive) throw err;
    return (await pca.acquireTokenPopup({ scopes: SCOPES })).accessToken;
  }
}

async function graph(path, { method = "GET", body, token } = {}) {
  token = token || (await getToken());
  const res = await fetch(path.startsWith("http") ? path : GRAPH + path, {
    method,
    headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error("Graph " + res.status + ": " + text.slice(0, 300));
  }
  return res.status === 204 ? null : res.json();
}

/** Loads every mail folder in the mailbox with its full path, e.g. "Inbox / Clients / Acme". */
async function loadAllFolders(token) {
  token = token || (await getToken());
  const all = [];
  const fetchLevel = async (url, parentPath) => {
    let next = url;
    while (next) {
      const d = await graph(next, { token });
      for (const f of d.value) {
        const path = parentPath ? parentPath + " / " + f.displayName : f.displayName;
        all.push({ id: f.id, name: f.displayName, path });
        if (f.childFolderCount > 0) queue.push([f.id, path]);
      }
      next = d["@odata.nextLink"] || null;
    }
  };
  const sel = "$top=250&$select=id,displayName,childFolderCount";
  const queue = [];
  await fetchLevel("/me/mailFolders?" + sel, "");
  // Walk the tree a few folders at a time
  while (queue.length) {
    const batch = queue.splice(0, 6);
    await Promise.all(
      batch.map(([id, path]) =>
        fetchLevel("/me/mailFolders/" + encodeURIComponent(id) + "/childFolders?" + sel, path)
      )
    );
  }
  all.sort((a, b) => a.path.localeCompare(b.path, undefined, { sensitivity: "base" }));
  return all;
}

let _sweeping = null;
/**
 * Looks through recent Sent Items for emails tagged with a folder choice and moves them there.
 * Returns [{subject, folder}] for each email moved.
 */
function sweepSentItems() {
  if (_sweeping) return _sweeping; // never run two sweeps at once
  _sweeping = (async () => {
    const token = await getToken();
    const expand = encodeURIComponent("singleValueExtendedProperties($filter=id eq '" + CECP_PROP_ID + "')");
    const data = await graph(
      "/me/mailFolders/sentitems/messages?$top=50&$select=id,subject&$expand=" + expand,
      { token }
    );
    const filed = [];
    for (const m of data.value) {
      const ep = (m.singleValueExtendedProperties || [])[0];
      if (!ep || !ep.value) continue;
      let props;
      try { props = JSON.parse(ep.value); } catch { continue; }
      if (!props.fileTo || props.fileTo === "none") continue;
      try {
        await graph("/me/messages/" + encodeURIComponent(m.id) + "/move", {
          method: "POST",
          body: { destinationId: props.fileTo },
          token,
        });
        filed.push({ subject: m.subject || "(no subject)", folder: props.fileToName || "folder" });
      } catch (err) {
        console.warn("Send & File: couldn't move", m.subject, err);
      }
    }
    return filed;
  })().finally(() => { _sweeping = null; });
  return _sweeping;
}
