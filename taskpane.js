/* Send & File - task pane. Works in two modes:
 *  - Writing an email: pick where it should go once sent (stored on the email itself).
 *  - Reading an email: move it to a folder now. Pin the pane and it also files sent mail every minute.
 */

const $ = (id) => document.getElementById(id);
const state = { folders: [], byId: new Map(), wellKnown: new Set(), compose: false, people: [], convId: null };

// ---------- small helpers ----------
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const call = (fn) => new Promise((resolve, reject) =>
  fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))));
const show = (id, on = true) => $(id).classList.toggle("hidden", !on);
function setStatus(text, kind = "ok") {
  $("status").innerHTML = text ? `<div class="msg ${kind}">${esc(text)}</div>` : "";
}

// ---------- remembered choices (roam with your mailbox) ----------
const rs = () => Office.context.roamingSettings;
const getRecent = () => rs().get("recent") || [];
const getPeople = () => rs().get("people") || {};
function remember(folder) {
  const recent = [folder.id, ...getRecent().filter((id) => id !== folder.id)].slice(0, 8);
  rs().set("recent", recent);
  const people = getPeople();
  for (const email of state.people) people[email] = folder.id;
  const keys = Object.keys(people);
  for (const k of keys.slice(0, Math.max(0, keys.length - 80))) delete people[k]; // keep it small
  rs().set("people", people);
  rs().saveAsync(() => {});
}

// ---------- rendering ----------
function folderButton(f, why) {
  const i = f.path.lastIndexOf(" / ");
  const parent = i > -1 ? f.path.slice(0, i) : "";
  return `<li><button data-id="${esc(f.id)}">
    <span class="fname">${esc(f.name)}</span>
    ${parent ? `<span class="fparent">${esc(parent)}</span>` : ""}
    ${why ? `<span class="why">${esc(why)}</span>` : ""}
  </button></li>`;
}
function renderList(ul, items) {
  $(ul).innerHTML = items.map(([f, why]) => folderButton(f, why)).join("");
}
function renderDest(id, path) {
  const d = $("dest");
  if (!state.compose) return;
  if (!id) {
    d.className = "dest unset"; $("destTab").textContent = "Not filed yet";
    $("destPath").textContent = "Pick a folder below"; $("destHint").textContent = "";
  } else if (id === "none") {
    d.className = "dest unset"; $("destTab").textContent = "Not filing";
    $("destPath").textContent = "Stays in Sent Items"; $("destHint").textContent = "You can press Send now.";
  } else {
    d.className = "dest"; $("destTab").textContent = "Files to";
    $("destPath").textContent = path; $("destHint").textContent = "Press Send and it will be moved here.";
  }
}

function search(q) {
  const words = q.toLowerCase().split(/\s+/).filter(Boolean);
  if (!words.length) { $("results").innerHTML = ""; return []; }
  const hits = state.folders
    .filter((f) => words.every((w) => f.path.toLowerCase().includes(w)))
    .sort((a, b) => (a.name.toLowerCase().startsWith(words[0]) ? -1 : 0) - (b.name.toLowerCase().startsWith(words[0]) ? -1 : 0))
    .slice(0, 30);
  if (!hits.length) $("results").innerHTML = `<li class="empty">No folder matches "${esc(q)}".</li>`;
  else renderList("results", hits.map((f) => [f]));
  return hits;
}

async function buildSuggestions() {
  const out = [];
  const seen = new Set();
  const add = (id, why) => {
    const f = state.byId.get(id);
    if (f && !seen.has(id) && !state.wellKnown.has(id)) { seen.add(id); out.push([f, why]); }
  };
  // Where the rest of this conversation is already filed
  if (state.convId) {
    try {
      const filter = encodeURIComponent(`conversationId eq '${state.convId.replace(/'/g, "''")}'`);
      const d = await graph(`/me/messages?$filter=${filter}&$select=parentFolderId&$top=25`);
      for (const m of d.value) add(m.parentFolderId, "Where this conversation is filed");
    } catch (e) { console.warn("conversation lookup failed", e); }
  }
  // Where you last filed mail to/from these people
  const people = getPeople();
  for (const email of state.people) if (people[email]) add(people[email], "Last used for " + email);
  renderList("suggested", out.slice(0, 5));
  show("suggestBlock", out.length > 0);

  const recent = getRecent().map((id) => state.byId.get(id)).filter(Boolean).map((f) => [f]);
  renderList("recent", recent);
  show("recentBlock", recent.length > 0);
}

// ---------- actions ----------
async function chooseForCompose(folder) {
  const item = Office.context.mailbox.item;
  const id = folder ? folder.id : "none";
  const path = folder ? folder.path : "";
  // sessionData is what the Send check reads; custom properties travel with the sent copy.
  await call((cb) => item.sessionData.setAsync("fileTo", id, cb));
  await call((cb) => item.sessionData.setAsync("fileToName", path, cb));
  const props = await call((cb) => item.loadCustomPropertiesAsync(cb));
  props.set("fileTo", id);
  props.set("fileToName", path);
  await call((cb) => props.saveAsync(cb));
  renderDest(id, path);
  if (folder) remember(folder);
  setStatus("");
}

async function moveCurrentEmail(folder) {
  const item = Office.context.mailbox.item;
  if (!item || !item.itemId) { setStatus("Select an email first.", "err"); return; }
  const restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
  await graph(`/me/messages/${encodeURIComponent(restId)}/move`, { method: "POST", body: { destinationId: folder.id } });
  remember(folder);
  setStatus(`Moved to ${folder.path}.`);
}

async function pick(id) {
  const folder = state.byId.get(id);
  if (!folder) return;
  try {
    if (state.compose) await chooseForCompose(folder);
    else await moveCurrentEmail(folder);
    $("q").value = ""; $("results").innerHTML = "";
    buildSuggestions();
  } catch (e) {
    setStatus("That didn't work: " + (e.message || e), "err");
  }
}

// ---------- background filing of sent mail ----------
async function runSweep() {
  try {
    const filed = await sweepSentItems();
    if (filed.length) {
      $("sweepLog").insertAdjacentHTML("afterbegin",
        filed.map((f) => `<li>Filed "${esc(f.subject)}" to ${esc(f.folder)}</li>`).join(""));
    }
  } catch (e) {
    $("sweepState").textContent = "Couldn't check Sent Items just now. Will try again shortly.";
    console.warn(e);
  }
}

// ---------- reading the current email's context ----------
async function loadItemContext() {
  const item = Office.context.mailbox.item;
  state.people = [];
  state.convId = item ? item.conversationId || null : null;
  if (!item) return;
  if (state.compose) {
    for (const field of ["to", "cc"]) {
      try {
        const list = await call((cb) => item[field].getAsync(cb));
        state.people.push(...list.map((r) => r.emailAddress.toLowerCase()));
      } catch (e) { /* field unavailable */ }
    }
    const id = await call((cb) => item.sessionData.getAsync("fileTo", cb)).catch(() => null);
    const path = await call((cb) => item.sessionData.getAsync("fileToName", cb)).catch(() => "");
    renderDest(id, path);
  } else if (item.from) {
    state.people.push(item.from.emailAddress.toLowerCase());
  }
}

async function loadFolders() {
  show("loading", true);
  const token = await getToken();
  const [folders, ...wk] = await Promise.all([
    loadAllFolders(token),
    ...["inbox", "sentitems", "drafts", "deleteditems", "junkemail", "outbox"].map((n) =>
      graph(`/me/mailFolders/${n}?$select=id`, { token }).then((f) => f.id).catch(() => null)),
  ]);
  state.folders = folders;
  state.byId = new Map(folders.map((f) => [f.id, f]));
  state.wellKnown = new Set(wk.filter(Boolean));
  show("loading", false);
  show("main", true);
}

async function start() {
  await loadFolders();
  await loadItemContext();
  await buildSuggestions();
  runSweep();
  setInterval(runSweep, 60000);
  $("q").focus();
}

// ---------- wiring ----------
Office.onReady(async () => {
  const item = Office.context.mailbox.item;
  state.compose = !!(item && item.to && typeof item.to.getAsync === "function");

  if (!state.compose) {
    $("title").textContent = "Move this email";
    show("dest", false);
    show("composeActions", false);
    $("sweepState").textContent = "Pin this pane to keep your sent emails filed automatically.";
    // Pinned pane: refresh when you select a different email
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, async () => {
      setStatus("");
      await loadItemContext();
      buildSuggestions();
    });
  }

  document.body.addEventListener("click", (e) => {
    const b = e.target.closest("ul.folders button[data-id]");
    if (b) pick(b.dataset.id);
  });
  $("q").addEventListener("input", (e) => search(e.target.value));
  $("q").addEventListener("keydown", (e) => {
    if (e.key === "Enter") { const hits = search(e.target.value); if (hits[0]) pick(hits[0].id); }
  });
  $("noFileBtn").addEventListener("click", () =>
    chooseForCompose(null).catch((err) => setStatus("That didn't work: " + err.message, "err")));
  $("refreshBtn").addEventListener("click", async () => {
    await loadFolders(); await buildSuggestions();
  });
  $("signinBtn").addEventListener("click", async () => {
    try { await getToken(true); show("signin", false); await start(); }
    catch (e) { $("signinErr").innerHTML = `<div class="msg err">Sign-in didn't complete: ${esc(e.message || e)}</div>`; }
  });

  try {
    await getToken(false);
    await start();
  } catch (e) {
    console.warn("Silent sign-in failed", e);
    show("loading", false);
    show("signin", true);
  }
});
