/* Private Notes – stores the note as a custom property on the mail item / appointment
   in YOUR mailbox (not sent to recipients or attendees). */
(function () {
  "use strict";

  const KEY = "privateNote";
  const KEY_TS = "privateNoteSaved";
  const SAVE_DELAY_MS = 800;

  const $ = (id) => document.getElementById(id);
  const noteEl = $("note"), statusEl = $("status"), countEl = $("count"),
        clearBtn = $("clear"), subjectEl = $("subject");

  let props = null;      // CustomProperties of the currently loaded item
  let loadToken = 0;     // guards against out-of-order async loads
  let timer = null;

  Office.onReady(() => {
    noteEl.addEventListener("input", onInput);
    clearBtn.addEventListener("click", onClear);
    window.addEventListener("pagehide", flush);
    // Fires when the pane is pinned and the user selects another item.
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { flush(); loadItem(); });
    loadItem();
  });

  function setStatus(text) { statusEl.textContent = text || "\u00a0"; }
  function updateCount() { countEl.textContent = noteEl.value.length; }

  function showEmpty(on) {
    $("msg").hidden = !on;
    $("main").style.display = on ? "none" : "contents";
  }

  function loadItem() {
    const item = Office.context.mailbox.item;
    const token = ++loadToken;
    props = null;
    noteEl.disabled = true; clearBtn.disabled = true;

    if (!item) { showEmpty(true); return; }
    showEmpty(false);

    const subj = item.subject;
    subjectEl.textContent = typeof subj === "string" ? subj : "";
    if (subj && typeof subj === "object") {            // compose mode: subject is async
      subj.getAsync((r) => { if (token === loadToken && r.status === Office.AsyncResultStatus.Succeeded) subjectEl.textContent = r.value; });
    }

    item.loadCustomPropertiesAsync((result) => {
      if (token !== loadToken) return;                 // user already moved to another item
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        setStatus("Could not load note: " + result.error.message);
        return;
      }
      props = result.value;
      noteEl.value = props.get(KEY) || "";
      noteEl.disabled = false;
      clearBtn.disabled = !noteEl.value;
      updateCount();
      const ts = props.get(KEY_TS);
      setStatus(ts ? "Saved " + new Date(ts).toLocaleString() : "");
    });
  }

  function onInput() {
    updateCount();
    clearBtn.disabled = !noteEl.value;
    setStatus("Typing…");
    clearTimeout(timer);
    timer = setTimeout(save, SAVE_DELAY_MS);
  }

  function flush() {
    if (timer) { clearTimeout(timer); timer = null; save(); }
  }

  function save() {
    timer = null;
    if (!props) return;
    const text = noteEl.value;
    const target = props;                              // keep reference if the item changes mid-save
    if (text) {
      target.set(KEY, text);
      target.set(KEY_TS, new Date().toISOString());
    } else {
      target.remove(KEY);
      target.remove(KEY_TS);
    }
    target.saveAsync((r) => {
      if (target !== props) return;
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        setStatus(text ? "Saved " + new Date().toLocaleTimeString() : "Note deleted");
      } else {
        setStatus("Save failed: " + r.error.message);
      }
    });
  }

  function onClear() {
    if (!noteEl.value || !confirm("Delete this private note?")) return;
    noteEl.value = "";
    updateCount();
    clearBtn.disabled = true;
    clearTimeout(timer);
    save();
  }
})();
