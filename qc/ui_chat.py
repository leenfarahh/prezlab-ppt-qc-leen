"""The ask box, what it answers, and the one button that acts.

Rendered into the pages a designer is already on rather than given a page of its
own. A question about a deck arrives while they are looking at that deck, and
sending them somewhere else to ask it is the map problem this exists to solve.

The panel posts to /chat/{job_id} and draws what comes back: the sentence, the
links, and - when the designer asked for something to be done - the PLAN, in a
box with a button under it.

THE PLAN IS THE CONFIRMATION GATE, DRAWN. It says what would happen, to how
many records, on which slides, in a sentence the tool built from what actually
resolved rather than from what the model said. Nothing has happened when it
appears. Pressing the button posts to /chat/{job_id}/do, which performs it
through the same function the button on the page performs, and then reloads,
because a page showing counts from before a change is the one thing this must
not do.

So the copy says both halves plainly. A box that looks like it might have
already changed something invites a designer to stop reading it, and a box that
looks like it cannot do anything invites them to go and find the page.

Rendering only; route logic lives in qc/web.py.
"""

from .ui import esc

_CSS = """
.askbox { background: #fff; border: 1px solid var(--line-soft); border-radius: 12px;
  padding: 1.1rem 1.2rem; margin: 1.2rem 0; }
.askrow { display: flex; gap: 0.5rem; margin-top: 0.6rem; }
.askrow input { flex: 1; min-width: 10rem; border: 1px solid var(--line);
  border-radius: 8px; padding: 0.55rem 0.7rem; font: inherit; color: var(--teal); }
.askrow input::placeholder { color: var(--slate); }
.askans { margin-top: 0.9rem; font-weight: 500; color: var(--teal); }
.askans.thin { color: var(--slate-text); font-weight: 400; }
.asklinks { display: flex; flex-wrap: wrap; gap: 0.45rem; margin-top: 0.6rem; }
.askex { color: var(--slate-text); font-size: 0.8rem; margin-top: 0.5rem; }
.askex button { background: none; border: 0; padding: 0; font: inherit;
  color: var(--teal); text-decoration: underline; cursor: pointer; }
.askplan { margin-top: 0.8rem; background: var(--sand); border: 1px solid var(--line);
  border-radius: 10px; padding: 0.85rem 1rem; }
.askplan .what { color: var(--teal); font-weight: 500; }
.askplan .gate { color: var(--slate-text); font-size: 0.8rem;
  margin-top: 0.35rem; }
.askplan .row { display: flex; flex-wrap: wrap; gap: 0.45rem;
  margin-top: 0.7rem; align-items: center; }
.asknope { color: var(--slate-text); font-size: 0.85rem; margin-top: 0.6rem; }
"""

# Four questions that each land somewhere different, so the examples teach the
# surface as well as saving typing. The last one is an ASK rather than a
# question: a designer who never discovers that this can act will keep going to
# find the page, which is the thing it was built to stop.
_EXAMPLES = ("What is worth fixing first?",
             "Which colours were typed in by hand?",
             "Is this master missing a layout?",
             "Fix what you can on this deck")


def chat_css() -> str:
    """The panel's styles as a STYLE BLOCK. Callers drop it into a page body
    beside the page's own, so bare CSS here would render as text."""
    return f"<style>{_CSS}</style>"


def chat_panel(job_id: str, kind: str, available: bool,
               unavailable_note: str = "") -> str:
    """The ask box for one job. `available` is false when no model is configured
    on this host, and then the box is not drawn at all - a text field that
    silently does nothing is worse than an absence with a reason beside it."""
    if not available:
        if not unavailable_note:
            return ""
        return (f'<div class="askbox"><div class="tag">Ask about this deck</div>'
                f'<p class="note" style="margin:0.3rem 0 0">'
                f'{esc(unavailable_note)}</p></div>')

    examples = " &middot; ".join(
        f'<button type="button" class="askeg" data-q="{esc(q)}">{esc(q)}</button>'
        for q in _EXAMPLES)

    return f"""
<div class="askbox">
  <div class="tag">Ask about this deck</div>
  <h3 style="margin:0.2rem 0 0.1rem">What do you want to know, or done?</h3>
  <p class="note" style="margin:0">Answers come from what the passes on this
  deck actually recorded. Ask for something to be done and it comes back as a
  plan saying exactly what would change: nothing happens until you press the
  button on it. It cannot read the words on your slides.</p>
  <form class="askrow" id="askform-{esc(job_id)}" data-busy="off"
        data-job="{esc(job_id)}" data-kind="{esc(kind)}">
    <input type="text" name="q" id="askq-{esc(job_id)}" autocomplete="off"
           placeholder="Ask about the layouts and colours, or ask for a fix">
    <button class="btn ghost" type="submit">Ask</button>
  </form>
  <div class="askex">Try: {examples}</div>
  <div class="askans" id="askans-{esc(job_id)}" aria-live="polite" hidden></div>
  <div class="asknope" id="asknope-{esc(job_id)}" hidden></div>
  <div id="askplan-{esc(job_id)}"></div>
  <div class="asklinks" id="asklinks-{esc(job_id)}"></div>
</div>
<script>
(function() {{
  const job = {job_id!r};
  const form = document.getElementById('askform-' + job);
  if (!form) return;
  const box = document.getElementById('askq-' + job);
  const out = document.getElementById('askans-' + job);
  const nope = document.getElementById('asknope-' + job);
  const planbox = document.getElementById('askplan-' + job);
  const links = document.getElementById('asklinks-' + job);

  function say(text, thin) {{
    out.hidden = false;
    out.textContent = text;
    out.classList.toggle('thin', !!thin);
  }}
  function clear() {{
    links.innerHTML = '';
    planbox.innerHTML = '';
    nope.hidden = true;
    nope.textContent = '';
  }}

  // The plan, drawn. Built here rather than sent as markup so the server never
  // has to trust this panel with HTML, and so the gate sentence under it is
  // the same on every page.
  function drawPlan(plan) {{
    const wrap = document.createElement('div');
    wrap.className = 'askplan';
    const what = document.createElement('div');
    what.className = 'what';
    what.textContent = plan.summary;
    const gate = document.createElement('div');
    gate.className = 'gate';
    gate.textContent = plan.changes
      ? 'Nothing has changed yet. This is what would happen if you press the button.'
      : 'This changes nothing in the deck.';
    const row = document.createElement('div');
    row.className = 'row';
    const go = document.createElement('button');
    go.className = 'btn primary';
    go.type = 'button';
    go.textContent = plan.changes ? 'Do it' : 'Go ahead';
    const drop = document.createElement('button');
    drop.className = 'btn ghost';
    drop.type = 'button';
    drop.textContent = 'Not that';
    drop.addEventListener('click', function() {{ clear(); }});

    go.addEventListener('click', function() {{
      go.disabled = true;
      drop.disabled = true;
      showBusy('Doing it',
               'Applying the change and re-reading the deck, so the counts on '
               + 'this page come from the file rather than from an assumption.');
      fetch('/chat/' + job + '/do', {{
        method: 'POST',
        headers: {{'Content-Type': 'application/json'}},
        body: JSON.stringify({{token: plan.token}})
      }}).then(r => r.json()).then(function(d) {{
        if (d.error) {{
          hideBusy();
          go.disabled = false;
          drop.disabled = false;
          gate.textContent = d.error;
          return;
        }}
        // The deck has changed, so every count and every picture on this page
        // is now a description of a file that no longer exists. Reloading is
        // the only honest end to this.
        if (d.reload) {{ window.location.reload(); return; }}
        hideBusy();
        wrap.innerHTML = '';
        const done = document.createElement('div');
        done.className = 'what';
        done.textContent = d.note;
        wrap.appendChild(done);
      }}).catch(function() {{
        hideBusy();
        go.disabled = false;
        drop.disabled = false;
        gate.textContent = 'That could not be sent, so nothing was changed.';
      }});
    }});

    row.appendChild(go);
    row.appendChild(drop);
    wrap.appendChild(what);
    wrap.appendChild(gate);
    wrap.appendChild(row);
    planbox.appendChild(wrap);
  }}

  form.addEventListener('submit', function(e) {{
    e.preventDefault();
    const q = box.value.trim();
    if (!q) return;
    clear();
    say('Reading what the passes recorded...', true);
    fetch('/chat/' + job, {{
      method: 'POST',
      headers: {{'Content-Type': 'application/json'}},
      body: JSON.stringify({{q: q}})
    }}).then(r => r.json()).then(d => {{
      if (d.error) {{ say(d.error, true); return; }}
      say(d.answer, false);
      if (d.refusal) {{ nope.hidden = false; nope.textContent = d.refusal; }}
      if (d.plan) {{ drawPlan(d.plan); }}
      (d.links || []).forEach(function(l) {{
        const a = document.createElement('a');
        a.className = 'btn ghost';
        a.href = l.href;
        a.textContent = l.label;
        links.appendChild(a);
      }});
    }}).catch(function() {{
      say('That could not be sent. Nothing was changed either way.', true);
    }});
  }});

  document.querySelectorAll('.askeg').forEach(function(b) {{
    b.addEventListener('click', function() {{
      box.value = b.dataset.q;
      form.requestSubmit();
    }});
  }});
}})();
</script>"""
