/// <reference types="@kitajs/html" />
const EXAMPLE_PROMPTS = [
  "Create a Q1 sales report with revenue by region and a chart",
  "Fix all typos and tighten the executive summary",
  "Add a new slide summarizing the key metrics",
  "Reformat the pricing table and add a totals row",
];

export function Page({ currentSid }: { currentSid: string }) {
  return (
    <html lang="en">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>offidized</title>
        <script src="https://unpkg.com/htmx.org@2.0.4"></script>
        <script src="https://cdn.tailwindcss.com"></script>
        <style>{`
          .ai-out.streaming::after {
            content: '';
            display: inline-block;
            width: 4px; height: 14px;
            background: #52525b;
            animation: blink .8s step-end infinite;
            vertical-align: text-bottom;
            margin-left: 1px;
          }
          @keyframes blink { 50% { opacity: 0 } }
          .ai-out[data-stream]::before {
            content: 'thinking...';
            color: #52525b;
            font-size: 11px;
            font-style: italic;
            animation: pulse 1.5s ease-in-out infinite;
          }
          @keyframes pulse { 50% { opacity: .4 } }
          .ai-seg { display: block; }

          /* ── Tool blocks ── */
          .tool-block {
            font-family: ui-monospace, SFMono-Regular, monospace;
            font-size: 11px;
            margin: 4px 0;
            border: 1px solid rgba(63,63,70,.4);
            border-radius: 4px;
            overflow: hidden;
          }
          .tool-head {
            display: flex;
            align-items: baseline;
            gap: 8px;
            padding: 5px 10px;
            background: rgba(24,24,27,.8);
          }
          .tool-name {
            font-size: 10px;
            text-transform: uppercase;
            letter-spacing: .05em;
            color: #52525b;
            flex-shrink: 0;
          }
          .tool-inv {
            color: #a1a1aa;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            flex: 1;
          }
          .tool-inv .op { color: #71717a; }
          .tool-inv .ph { color: #a1a1aa; }
          .tool-inv .hi { color: #52525b; }
          .tool-out {
            padding: 6px 10px;
            color: #71717a;
            white-space: pre-wrap;
            word-break: break-word;
            line-height: 1.5;
            max-height: 180px;
            overflow-y: auto;
            border-top: 1px solid rgba(63,63,70,.3);
          }
          .tool-out.err { color: #fca5a5; }
          .tool-out.muted { color: #71717a; font-style: italic; }
          .tool-out.clipped { max-height: 88px; }
          .tool-expand {
            display: block;
            padding: 2px 10px 4px;
            color: #52525b;
            cursor: pointer;
            font-size: 10px;
            border-top: 1px solid rgba(63,63,70,.2);
            user-select: none;
          }
          .tool-expand:hover { color: #a1a1aa; }

          /* ── Turn-end marker ── */
          .turn-end {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 8px 0 2px;
            color: #3f3f46;
            font-size: 10px;
            font-family: ui-monospace, monospace;
          }
          .turn-end::before, .turn-end::after {
            content: '';
            flex: 1;
            height: 1px;
            background: rgba(63,63,70,.25);
          }

          /* ── Chat form states ── */
          #messages { scroll-behavior: smooth; }
          #chat-form.busy input,
          #chat-form.busy button { pointer-events: none; opacity: .5; }
          #chat-form.busy #abort-btn { pointer-events: auto; opacity: 1; }

          /* ── API key modal ── */
          #key-modal {
            display: none;
            position: fixed;
            inset: 0;
            z-index: 100;
            background: rgba(0,0,0,.75);
            align-items: center;
            justify-content: center;
          }
          #key-modal.open { display: flex; }
          #key-modal-box {
            background: #18181b;
            border: 1px solid rgba(63,63,70,.6);
            border-radius: 8px;
            padding: 20px 24px;
            width: 360px;
            max-width: 90vw;
          }

          /* ── Preview panel ── */
          #preview-container { background: #f5f5f5; }
          #preview-container.type-xlsx { background: #fff; }
          #preview-loading {
            position: absolute;
            inset: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #a1a1aa;
            font-size: 12px;
            z-index: 5;
          }
          #preview-loading.hidden { display: none; }

          /* ── Resize handle ── */
          #resize-handle {
            width: 5px;
            cursor: col-resize;
            background: transparent;
            transition: background .15s;
            flex-shrink: 0;
          }
          #resize-handle:hover,
          #resize-handle.active { background: rgba(63,63,70,.5); }

          /* ── Versions dropdown ── */
          #versions-dropdown {
            position: absolute;
            right: 0;
            top: calc(100% + 4px);
            background: #18181b;
            border: 1px solid rgba(63,63,70,.6);
            border-radius: 6px;
            box-shadow: 0 8px 24px rgba(0,0,0,.5);
            z-index: 50;
            min-width: 180px;
            max-height: 260px;
            overflow-y: auto;
            padding: 4px 0;
          }
          .ver-item {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 12px;
            padding: 6px 12px;
            color: #a1a1aa;
            text-decoration: none;
            font-size: 11px;
            white-space: nowrap;
            cursor: pointer;
          }
          .ver-item:hover { background: #27272a; color: #e4e4e7; }
          .ver-item.latest { color: #d4d4d8; }
          .ver-item.original { color: #a1a1aa; font-style: italic; }

          /* ── Sidebar ── */
          #sidebar {
            width: 220px;
            min-width: 220px;
            transition: margin-left .15s ease;
          }
          #sidebar.collapsed {
            margin-left: -220px;
          }
          .sess-item {
            display: block;
            padding: 8px 12px;
            cursor: pointer;
            border-radius: 4px;
            margin: 1px 6px;
            transition: background .1s;
          }
          .sess-item:hover { background: rgba(39,39,42,.8); }
          .sess-item.active { background: rgba(39,39,42,1); }
          .sess-title {
            font-size: 12px;
            color: #d4d4d8;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
          }
          .sess-meta {
            font-size: 10px;
            color: #52525b;
            margin-top: 1px;
          }
        `}</style>
      </head>
      <body class="bg-[#0c0c0c] text-zinc-300 h-screen flex flex-col text-[13px] antialiased">
        {/* ── API key modal ── */}
        <div id="key-modal" onclick="if(event.target===this)closeKeyModal()">
          <div id="key-modal-box">
            <div class="text-zinc-200 text-sm font-medium mb-2">API Keys</div>
            <div class="text-zinc-500 text-xs mb-3">
              Stored in browser only. First key found is used.
            </div>
            <div class="mb-2">
              <label class="text-zinc-400 text-[10px] uppercase tracking-wide">
                OpenRouter
              </label>
              <input
                id="openrouter-key-input"
                type="password"
                placeholder="sk-or-..."
                class="w-full bg-zinc-900 border border-zinc-700 rounded px-3 py-2 text-zinc-200 text-xs outline-none focus:border-zinc-500 mt-1"
              />
            </div>
            <div class="mb-2">
              <label class="text-zinc-400 text-[10px] uppercase tracking-wide">
                Google (Gemini)
              </label>
              <input
                id="google-key-input"
                type="password"
                placeholder="AIza..."
                class="w-full bg-zinc-900 border border-zinc-700 rounded px-3 py-2 text-zinc-200 text-xs outline-none focus:border-zinc-500 mt-1"
              />
            </div>
            <div class="mb-3">
              <label class="text-zinc-400 text-[10px] uppercase tracking-wide">
                Anthropic (Claude)
              </label>
              <input
                id="claude-key-input"
                type="password"
                placeholder="sk-ant-..."
                class="w-full bg-zinc-900 border border-zinc-700 rounded px-3 py-2 text-zinc-200 text-xs outline-none focus:border-zinc-500 mt-1"
              />
            </div>
            <div class="flex gap-2 justify-end">
              <button
                onclick="closeKeyModal()"
                class="px-3 py-1.5 text-xs text-zinc-500 hover:text-zinc-300 transition-colors"
              >
                Cancel
              </button>
              <button
                onclick="saveKeys()"
                class="px-3 py-1.5 text-xs bg-zinc-800 hover:bg-zinc-700 text-zinc-200 rounded transition-colors"
              >
                Save
              </button>
            </div>
          </div>
        </div>

        {/* ── Header ── */}
        <header class="shrink-0 h-11 flex items-center justify-between px-4 border-b border-zinc-800/60">
          <div class="flex items-center gap-2">
            <button
              id="sidebar-toggle"
              onclick="toggleSidebar()"
              title="Toggle sidebar"
              class="w-8 h-8 flex items-center justify-center rounded text-zinc-600 hover:text-zinc-300 hover:bg-zinc-800 transition-colors"
            >
              <svg
                width="14"
                height="14"
                viewBox="0 0 15 15"
                fill="none"
                stroke="currentColor"
                stroke-width="1.5"
                stroke-linecap="round"
                stroke-linejoin="round"
              >
                <path d="M2 3.5h11M2 7.5h11M2 11.5h11" />
              </svg>
            </button>
            <span class="text-[13px] font-semibold text-zinc-100 tracking-tight">
              offidized
            </span>
          </div>
          <div class="flex items-center gap-1">
            {/* file indicator */}
            <div
              id="file-indicator"
              class="hidden items-center gap-2 text-xs text-zinc-500 mr-2"
            >
              <span id="file-tag"></span>
              <div class="relative" id="versions-wrap">
                <button
                  id="dl-btn"
                  onclick="toggleVersions(event)"
                  class="text-zinc-500 hover:text-zinc-200 transition-colors flex items-center"
                  title="Download versions"
                >
                  <svg
                    width="13"
                    height="13"
                    viewBox="0 0 15 15"
                    fill="none"
                    stroke="currentColor"
                    stroke-width="1.5"
                    stroke-linecap="round"
                    stroke-linejoin="round"
                  >
                    <path d="M7.5 1v10M3.5 8l4 4 4-4M2 13h11" />
                  </svg>
                </button>
                <div id="versions-dropdown" class="hidden"></div>
              </div>
              <button
                onclick="removeFile()"
                class="text-zinc-600 hover:text-zinc-300"
              >
                &times;
              </button>
            </div>

            {/* api key button */}
            <button
              onclick="openKeyModal()"
              title="Set API keys (Google/Claude)"
              class="relative w-8 h-8 flex items-center justify-center rounded text-zinc-600 hover:text-zinc-300 hover:bg-zinc-800 transition-colors"
            >
              <svg
                width="14"
                height="14"
                viewBox="0 0 15 15"
                fill="none"
                stroke="currentColor"
                stroke-width="1.5"
                stroke-linecap="round"
                stroke-linejoin="round"
              >
                <path d="M5 7a4 4 0 108 0A4 4 0 005 7zM5 7H1m3-2-2-2m2 6-2 2" />
              </svg>
              <span
                id="key-dot"
                class="hidden absolute top-1.5 right-1.5 w-1.5 h-1.5 rounded-full bg-emerald-500"
              ></span>
            </button>
          </div>
        </header>

        {/* ── Main: sidebar + chat + preview ── */}
        <div id="outer" class="flex-1 flex overflow-hidden min-h-0">
          {/* ── Sidebar ── */}
          <div
            id="sidebar"
            class="flex flex-col border-r border-zinc-800/60 bg-[#0a0a0a] overflow-hidden"
          >
            <div class="shrink-0 p-2">
              <button
                onclick="newChat()"
                class="w-full flex items-center gap-2 px-3 py-2 text-xs text-zinc-400 hover:text-zinc-200 bg-zinc-900/60 hover:bg-zinc-800 rounded border border-zinc-800/50 transition-colors"
              >
                <svg
                  width="12"
                  height="12"
                  viewBox="0 0 15 15"
                  fill="none"
                  stroke="currentColor"
                  stroke-width="1.5"
                  stroke-linecap="round"
                  stroke-linejoin="round"
                >
                  <path d="M7.5 2v11M2 7.5h11" />
                </svg>
                New chat
              </button>
            </div>
            <div id="session-list" class="flex-1 overflow-y-auto pb-2"></div>
          </div>

          {/* chat + preview area */}
          <div id="main" class="flex-1 flex overflow-hidden min-h-0">
            {/* chat panel */}
            <div
              id="chat-panel"
              class="flex flex-col min-w-0"
              style="flex: 1 1 0%"
            >
              {/* messages */}
              <div id="messages" class="flex-1 overflow-y-auto">
                <div
                  id="empty"
                  class="h-full flex flex-col items-center justify-center gap-6 text-zinc-600 select-none px-8"
                >
                  <div class="text-center">
                    <div class="text-xs tracking-widest uppercase text-zinc-600 mb-2">
                      offidized
                    </div>
                    <div class="text-[13px] text-zinc-500 leading-relaxed max-w-xs">
                      AI that edits Office files without destroying your
                      formatting, formulas, or charts.
                    </div>
                  </div>
                  <div class="w-full max-w-xs space-y-1.5">
                    {EXAMPLE_PROMPTS.map((p) => (
                      <button
                        onclick={`fillPrompt(${JSON.stringify(p)})`}
                        class="w-full text-left text-[12px] text-zinc-600 hover:text-zinc-300 bg-zinc-900/60 hover:bg-zinc-800 rounded px-3 py-2 transition-colors border border-zinc-800/50"
                      >
                        {p}
                      </button>
                    ))}
                  </div>
                  <div class="text-[11px] text-zinc-700">
                    Drop a file to edit, or describe one to create from scratch.
                  </div>
                </div>
              </div>

              {/* input */}
              <div class="shrink-0 border-t border-zinc-800/60">
                <form
                  id="chat-form"
                  hx-post="/api/chat"
                  hx-target="#messages"
                  hx-swap="beforeend"
                  hx-on--after-swap="document.getElementById('fd').value='';document.getElementById('empty')?.remove();document.getElementById('messages').scrollTo({top:99999})"
                  class="flex items-center gap-0 px-3 py-2"
                >
                  <input type="hidden" name="filename" id="fn" value="" />
                  <input type="hidden" name="filedata" id="fd" value="" />

                  <button
                    type="button"
                    onclick="document.getElementById('file-pick').click()"
                    class="shrink-0 w-8 h-8 flex items-center justify-center rounded text-zinc-500 hover:text-zinc-200 hover:bg-zinc-800 transition-colors"
                    title="Upload file"
                  >
                    <svg
                      width="15"
                      height="15"
                      viewBox="0 0 15 15"
                      fill="none"
                      stroke="currentColor"
                      stroke-width="1.5"
                      stroke-linecap="round"
                      stroke-linejoin="round"
                    >
                      <path d="M13.5 6.5l-5.59-5.3a3.18 3.18 0 00-4.32.12 3.18 3.18 0 00.12 4.32l6.3 5.96a2.12 2.12 0 002.88-.08 2.12 2.12 0 00-.08-2.88L6.5 3.07a1.06 1.06 0 00-1.44.04 1.06 1.06 0 00.04 1.44l4.95 4.68" />
                    </svg>
                  </button>
                  <input
                    id="file-pick"
                    type="file"
                    accept=".xlsx,.docx,.pptx"
                    class="hidden"
                  />

                  <textarea
                    name="message"
                    id="msg-input"
                    placeholder="Describe a document..."
                    class="flex-1 bg-transparent px-2 py-1.5 text-zinc-200 placeholder-zinc-600 outline-none resize-none"
                    required
                    rows="1"
                  ></textarea>

                  {/* abort button — shown while streaming */}
                  <button
                    id="abort-btn"
                    type="button"
                    class="hidden shrink-0 w-8 h-8 items-center justify-center rounded text-zinc-500 hover:text-red-400 hover:bg-zinc-800 transition-colors"
                    title="Stop generation"
                  >
                    <svg
                      width="11"
                      height="11"
                      viewBox="0 0 11 11"
                      fill="currentColor"
                    >
                      <rect x="1" y="1" width="9" height="9" rx="2" />
                    </svg>
                  </button>

                  <button
                    type="submit"
                    id="send-btn"
                    aria-label="Send"
                    class="shrink-0 w-8 h-8 flex items-center justify-center rounded text-zinc-500 hover:text-zinc-200 hover:bg-zinc-800 transition-colors"
                  >
                    <svg
                      width="15"
                      height="15"
                      viewBox="0 0 15 15"
                      fill="none"
                      stroke="currentColor"
                      stroke-width="1.5"
                      stroke-linecap="round"
                      stroke-linejoin="round"
                    >
                      <path d="M14 1L7 8M14 1l-4.5 13-2-5.5L2 4z" />
                    </svg>
                  </button>
                </form>
              </div>
            </div>

            {/* resize handle */}
            <div id="resize-handle" class="hidden"></div>

            {/* preview panel */}
            <div
              id="preview-panel"
              class="hidden flex-col border-l border-zinc-800/60 min-w-0"
              style="flex: 1 1 0%"
            >
              <div class="shrink-0 h-8 flex items-center px-3 text-[11px] text-zinc-500 border-b border-zinc-800/40 gap-2">
                <span id="preview-title">Preview</span>
                <span class="flex-1"></span>
                <span id="preview-time" class="text-zinc-600"></span>
              </div>
              <div
                id="preview-container"
                class="flex-1 relative overflow-hidden"
              >
                <div id="preview-loading" class="hidden">
                  Loading viewer...
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* drop overlay */}
        <div
          id="drop-overlay"
          class="hidden fixed inset-0 z-50 bg-[#0c0c0c]/90 flex items-center justify-center"
        >
          <div class="border border-dashed border-zinc-700 rounded-lg px-10 py-6 text-zinc-500 text-sm">
            Drop .xlsx, .docx, or .pptx
          </div>
        </div>

        {/* ── Main script ── */}
        <script>{`
          var currentSid = ${JSON.stringify(currentSid)};
          var filePick = document.getElementById('file-pick');
          var fnInput  = document.getElementById('fn');
          var fdInput  = document.getElementById('fd');
          var fileTag  = document.getElementById('file-tag');
          var fileInd  = document.getElementById('file-indicator');
          var overlay  = document.getElementById('drop-overlay');
          var msgs     = document.getElementById('messages');
          var form     = document.getElementById('chat-form');
          var msgInput = document.getElementById('msg-input');
          var abortBtn = document.getElementById('abort-btn');
          var sidebar  = document.getElementById('sidebar');
          var dragDepth = 0;
          var activeEs = null;

          // ── Textarea: Enter to send, Shift+Enter for newline, auto-resize ──
          msgInput.addEventListener('keydown', function(e) {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault();
              if (msgInput.value.trim() && !form.classList.contains('busy')) {
                htmx.trigger(form, 'submit');
              }
            }
          });
          msgInput.addEventListener('input', function() {
            this.style.height = 'auto';
            this.style.height = Math.min(this.scrollHeight, 120) + 'px';
          });

          // ── Prevent double-submit: go busy + clear input immediately ──
          document.body.addEventListener('htmx:beforeRequest', function(evt) {
            if (evt.detail.elt === form) {
              form.classList.add('busy');
              msgInput.value = '';
              msgInput.style.height = '';
            }
          });

          // Inject API keys into every HTMX request
          document.body.addEventListener('htmx:configRequest', function(evt) {
            var openrouterKey = sessionStorage.getItem('openrouterKey');
            var googleKey = sessionStorage.getItem('googleKey');
            var claudeKey = sessionStorage.getItem('claudeKey');
            if (openrouterKey) evt.detail.headers['X-OpenRouter-Key'] = openrouterKey;
            if (googleKey) evt.detail.headers['X-Google-Key'] = googleKey;
            if (claudeKey) evt.detail.headers['X-Claude-Key'] = claudeKey;
          });

          function fillPrompt(text) {
            msgInput.value = text;
            msgInput.focus();
          }

          // ── Sidebar ──
          function toggleSidebar() {
            sidebar.classList.toggle('collapsed');
          }

          function timeAgo(ts) {
            var diff = Math.floor((Date.now() - ts) / 1000);
            if (diff < 60) return 'just now';
            if (diff < 3600) return Math.floor(diff / 60) + 'm ago';
            if (diff < 86400) return Math.floor(diff / 3600) + 'h ago';
            return Math.floor(diff / 86400) + 'd ago';
          }

          async function refreshSidebar() {
            try {
              var res = await fetch('/api/sessions');
              if (!res.ok) return;
              var sessions = await res.json();
              var list = document.getElementById('session-list');
              if (!sessions.length) {
                list.innerHTML = '<div class="px-4 py-6 text-center text-zinc-700 text-[11px]">No sessions yet</div>';
                return;
              }
              list.innerHTML = sessions.map(function(s) {
                var active = s.id === currentSid ? ' active' : '';
                var icon = s.hasFile ? '<span class="text-zinc-600 mr-1">\\u{1F4CE}</span>' : '';
                var count = s.messageCount ? '<span class="text-zinc-700"> \\xB7 ' + s.messageCount + ' msg' + (s.messageCount > 1 ? 's' : '') + '</span>' : '';
                return '<div class="sess-item' + active + '" onclick="switchSession(\\'' + s.id + '\\')" title="' + s.title.replace(/"/g, '&quot;') + '">' +
                  '<div class="sess-title">' + icon + s.title.replace(/</g, '&lt;') + '</div>' +
                  '<div class="sess-meta">' + timeAgo(s.lastActive) + count + '</div>' +
                  '</div>';
              }).join('');
            } catch {}
          }

          async function switchSession(id) {
            if (id === currentSid) return;
            if (activeEs) { activeEs.close(); activeEs = null; }
            try {
              var res = await fetch('/api/switch/' + id);
              if (!res.ok) return;
              var data = await res.json();
              currentSid = id;

              // Restore conversation
              if (data.html) {
                msgs.innerHTML = data.html;
              } else {
                showEmptyState();
              }

              // Restore file indicator
              if (data.currentFile && data.hasFile) {
                fnInput.value = data.currentFile;
                fileTag.textContent = data.currentFile;
                fileInd.style.display = 'flex';
                msgInput.placeholder = 'Ask me to edit ' + data.currentFile + '...';
                if (window._showPreviewFromServer) window._showPreviewFromServer(data.currentFile);
              } else {
                removeFile();
              }

              form.classList.remove('busy');
              abortBtn.classList.add('hidden');
              abortBtn.classList.remove('flex');
              refreshSidebar();
            } catch {}
          }

          var examplePrompts = ${JSON.stringify(EXAMPLE_PROMPTS)};
          function showEmptyState() {
            var prompts = examplePrompts.map(function(p) {
              return '<button onclick="fillPrompt(' + JSON.stringify(p).replace(/"/g, '&quot;') + ')" ' +
                'class="w-full text-left text-[12px] text-zinc-600 hover:text-zinc-300 bg-zinc-900/60 hover:bg-zinc-800 rounded px-3 py-2 transition-colors border border-zinc-800/50">' +
                p.replace(/</g, '&lt;') + '</button>';
            }).join('');
            msgs.innerHTML = '<div id="empty" class="h-full flex flex-col items-center justify-center gap-6 text-zinc-600 select-none px-8">' +
              '<div class="text-center"><div class="text-xs tracking-widest uppercase text-zinc-600 mb-2">offidized</div>' +
              '<div class="text-[13px] text-zinc-500 leading-relaxed max-w-xs">AI that edits Office files without destroying your formatting, formulas, or charts.</div></div>' +
              '<div class="w-full max-w-xs space-y-1.5">' + prompts + '</div>' +
              '<div class="text-[11px] text-zinc-700">Drop a file to edit, or describe one to create from scratch.</div></div>';
          }

          // ── API key modal ──
          function openKeyModal() {
            document.getElementById('openrouter-key-input').value = sessionStorage.getItem('openrouterKey') || '';
            document.getElementById('google-key-input').value = sessionStorage.getItem('googleKey') || '';
            document.getElementById('claude-key-input').value = sessionStorage.getItem('claudeKey') || '';
            document.getElementById('key-modal').classList.add('open');
            setTimeout(function() { document.getElementById('openrouter-key-input').focus(); }, 50);
          }
          function closeKeyModal() {
            document.getElementById('key-modal').classList.remove('open');
          }
          function saveKeys() {
            var openrouterVal = document.getElementById('openrouter-key-input').value.trim();
            var googleVal = document.getElementById('google-key-input').value.trim();
            var claudeVal = document.getElementById('claude-key-input').value.trim();
            if (openrouterVal) {
              sessionStorage.setItem('openrouterKey', openrouterVal);
            } else {
              sessionStorage.removeItem('openrouterKey');
            }
            if (googleVal) {
              sessionStorage.setItem('googleKey', googleVal);
            } else {
              sessionStorage.removeItem('googleKey');
            }
            if (claudeVal) {
              sessionStorage.setItem('claudeKey', claudeVal);
            } else {
              sessionStorage.removeItem('claudeKey');
            }
            updateKeyDot();
            closeKeyModal();
          }
          function updateKeyDot() {
            var hasOpenRouter = !!sessionStorage.getItem('openrouterKey');
            var hasGoogle = !!sessionStorage.getItem('googleKey');
            var hasClaude = !!sessionStorage.getItem('claudeKey');
            var dot = document.getElementById('key-dot');
            if (hasOpenRouter || hasGoogle || hasClaude) dot.classList.remove('hidden');
            else dot.classList.add('hidden');
          }
          document.getElementById('openrouter-key-input').addEventListener('keydown', function(e) {
            if (e.key === 'Enter') saveKeys();
            if (e.key === 'Escape') closeKeyModal();
          });
          document.getElementById('google-key-input').addEventListener('keydown', function(e) {
            if (e.key === 'Enter') saveKeys();
            if (e.key === 'Escape') closeKeyModal();
          });
          document.getElementById('claude-key-input').addEventListener('keydown', function(e) {
            if (e.key === 'Enter') saveKeys();
            if (e.key === 'Escape') closeKeyModal();
          });
          // Restore key dot on page load
          updateKeyDot();

          // ── File handling ──
          function handleFile(file) {
            var ext = file.name.split('.').pop().toLowerCase();
            if (['xlsx','docx','pptx'].indexOf(ext) === -1) return;
            fnInput.value = file.name;
            fileTag.textContent = file.name;
            fileInd.style.display = 'flex';
            msgInput.placeholder = 'Ask me to edit ' + file.name + '...';
            var r = new FileReader();
            r.onload = function() {
              var bytes = new Uint8Array(r.result);
              var bin = '';
              for (var i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
              fdInput.value = btoa(bin);
              if (window._showPreview) window._showPreview(r.result, file.name);
            };
            r.readAsArrayBuffer(file);
          }

          function removeFile() {
            fnInput.value = '';
            fdInput.value = '';
            fileTag.textContent = '';
            fileInd.style.display = 'none';
            filePick.value = '';
            msgInput.placeholder = 'Describe a document...';
            closeVersions();
            if (window._hidePreview) window._hidePreview();
          }

          // ── Versions dropdown ──
          var versionsOpen = false;
          function toggleVersions(e) {
            e.stopPropagation();
            if (versionsOpen) { closeVersions(); return; }
            var drop = document.getElementById('versions-dropdown');
            drop.classList.remove('hidden');
            versionsOpen = true;
            loadVersions();
          }
          function closeVersions() {
            document.getElementById('versions-dropdown').classList.add('hidden');
            versionsOpen = false;
          }
          async function loadVersions() {
            var drop = document.getElementById('versions-dropdown');
            if (!drop) return;
            drop.innerHTML = '<div class="ver-item" style="color:#52525b;font-style:italic">Loading\\u2026</div>';
            try {
              var res = await fetch('/api/versions');
              if (!res.ok) { drop.innerHTML = ''; return; }
              var vers = await res.json();
              if (!vers.length) {
                drop.innerHTML = '<div class="ver-item" style="color:#52525b;font-style:italic">No versions yet</div>';
                return;
              }
              var dl = '<svg width="10" height="10" viewBox="0 0 15 15" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M7.5 1v10M3.5 8l4 4 4-4M2 13h11"/></svg>';
              drop.innerHTML = vers.map(function(v, i) {
                var tag = v.isOriginal ? 'Original' : (i === 0 ? 'Latest' : 'v' + v.n);
                var label = tag + ' \\xB7 ' + v.label;
                var cls = v.isOriginal ? ' original' : (i === 0 ? ' latest' : '');
                return '<a href="/api/download/v/' + v.n + '" download class="ver-item' + cls + '">' +
                  '<span>' + label + '</span>' + dl + '</a>';
              }).join('');
            } catch { drop.innerHTML = ''; }
          }
          document.addEventListener('click', function(e) {
            var wrap = document.getElementById('versions-wrap');
            if (versionsOpen && wrap && !wrap.contains(e.target)) closeVersions();
          });
          window._reloadVersions = function() { if (versionsOpen) loadVersions(); };

          filePick.addEventListener('change', function() {
            if (filePick.files[0]) handleFile(filePick.files[0]);
          });

          document.addEventListener('dragenter', function(e) {
            e.preventDefault();
            if (++dragDepth === 1) overlay.classList.remove('hidden');
          });
          document.addEventListener('dragleave', function(e) {
            e.preventDefault();
            if (--dragDepth === 0) overlay.classList.add('hidden');
          });
          document.addEventListener('dragover', function(e) { e.preventDefault(); });
          document.addEventListener('drop', function(e) {
            e.preventDefault();
            dragDepth = 0;
            overlay.classList.add('hidden');
            if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
          });

          // ── New chat ──
          async function newChat() {
            if (activeEs) { activeEs.close(); activeEs = null; }
            closeVersions();
            try {
              var res = await fetch('/api/reset', { method: 'POST' });
              var data = await res.json();
              currentSid = data.sid;
              removeFile();
              showEmptyState();
              form.classList.remove('busy');
              abortBtn.classList.add('hidden');
              abortBtn.classList.remove('flex');
              refreshSidebar();
            } catch {}
          }

          // ── Abort ──
          abortBtn.addEventListener('click', function() {
            if (activeEs) { activeEs.close(); activeEs = null; }
            form.classList.remove('busy');
            abortBtn.classList.add('hidden');
            abortBtn.classList.remove('flex');
          });

          // ── SSE streaming ──
          document.body.addEventListener('htmx:afterSettle', function() {
            var el = msgs.querySelector('[data-stream]');
            if (!el) return;
            var url = el.getAttribute('data-stream');
            el.removeAttribute('data-stream');
            el.classList.add('streaming');
            form.classList.add('busy');
            abortBtn.classList.remove('hidden');
            abortBtn.classList.add('flex');

            var es = new EventSource(url);
            activeEs = es;

            es.addEventListener('token', function(ev) {
              el.insertAdjacentHTML('beforeend', ev.data);
              msgs.scrollTo({ top: msgs.scrollHeight });
            });
            es.addEventListener('file-updated', function() {
              if (window._refreshPreview) window._refreshPreview();
            });
            es.addEventListener('file-created', function(ev) {
              var name = ev.data;
              if (!name) return;
              fnInput.value = name;
              fileTag.textContent = name;
              fileInd.style.display = 'flex';
              msgInput.placeholder = 'Ask me to edit ' + name + '...';
              if (window._showPreviewFromServer) window._showPreviewFromServer(name);
            });
            es.addEventListener('versions-updated', function() {
              if (window._reloadVersions) window._reloadVersions();
            });
            es.addEventListener('session-updated', function() {
              refreshSidebar();
            });
            es.addEventListener('done', function() {
              es.close(); activeEs = null;
              el.classList.remove('streaming');
              form.classList.remove('busy');
              abortBtn.classList.add('hidden');
              abortBtn.classList.remove('flex');
              refreshSidebar();
            });
            es.onerror = function() {
              es.close(); activeEs = null;
              el.classList.remove('streaming');
              form.classList.remove('busy');
              abortBtn.classList.add('hidden');
              abortBtn.classList.remove('flex');
            };
          });

          // ── Resize handle ──
          (function() {
            var handle = document.getElementById('resize-handle');
            var chatPanel = document.getElementById('chat-panel');
            var previewPanel = document.getElementById('preview-panel');
            var main = document.getElementById('main');
            var dragging = false;

            handle.addEventListener('mousedown', function(e) {
              e.preventDefault();
              dragging = true;
              handle.classList.add('active');
              document.body.style.cursor = 'col-resize';
              document.body.style.userSelect = 'none';
            });
            document.addEventListener('mousemove', function(e) {
              if (!dragging) return;
              var rect = main.getBoundingClientRect();
              var pct = ((e.clientX - rect.left) / rect.width) * 100;
              pct = Math.max(25, Math.min(75, pct));
              chatPanel.style.flex = '0 0 ' + pct + '%';
              previewPanel.style.flex = '0 0 ' + (100 - pct) + '%';
            });
            document.addEventListener('mouseup', function() {
              if (!dragging) return;
              dragging = false;
              handle.classList.remove('active');
              document.body.style.cursor = '';
              document.body.style.userSelect = '';
            });
          })();

          // ── Init: restore session + sidebar ──
          (async function() {
            try {
              var res = await fetch('/api/switch/' + currentSid);
              if (!res.ok) return;
              var data = await res.json();
              if (data.html) {
                msgs.innerHTML = data.html;
                msgs.scrollTo({ top: msgs.scrollHeight });
              }
              if (data.currentFile && data.hasFile) {
                fnInput.value = data.currentFile;
                fileTag.textContent = data.currentFile;
                fileInd.style.display = 'flex';
                msgInput.placeholder = 'Ask me to edit ' + data.currentFile + '...';
                if (window._showPreviewFromServer) window._showPreviewFromServer(data.currentFile);
              }
            } catch {}
            refreshSidebar();
          })();
        `}</script>

        {/* ── Preview module script ── */}
        <script type="module">{`
          var previewPanel = document.getElementById('preview-panel');
          var previewContainer = document.getElementById('preview-container');
          var previewTitle = document.getElementById('preview-title');
          var previewTime = document.getElementById('preview-time');
          var previewLoading = document.getElementById('preview-loading');
          var resizeHandle = document.getElementById('resize-handle');

          var viewer = null;
          var viewerType = null;
          var modules = {};

          async function getModule(ext) {
            if (modules[ext]) return modules[ext];
            var path = ext === 'xlsx' ? '/preview/xl-view.js'
                     : ext === 'docx' ? '/preview/doc-view.js'
                     : '/preview/ppt-view.js';
            modules[ext] = await import(path);
            return modules[ext];
          }

          function wasmUrl(ext) {
            return ext === 'xlsx' ? '/preview/xlview_bg.wasm'
                 : ext === 'docx' ? '/preview/offidized_docview_bg.wasm'
                 : '/preview/offidized_pptview_bg.wasm';
          }

          async function showPreview(arrayBuffer, filename) {
            var ext = filename.split('.').pop().toLowerCase();
            if (['xlsx','docx','pptx'].indexOf(ext) === -1) return;

            previewPanel.classList.remove('hidden');
            previewPanel.classList.add('flex');
            resizeHandle.classList.remove('hidden');
            previewTitle.textContent = filename;
            previewContainer.classList.toggle('type-xlsx', ext === 'xlsx');

            if (ext !== viewerType) {
              if (viewer) { try { viewer.destroy(); } catch {} viewer = null; }
              while (previewContainer.firstChild && previewContainer.firstChild !== previewLoading) {
                previewContainer.removeChild(previewContainer.firstChild);
              }
              if (previewLoading.parentNode !== previewContainer) {
                previewContainer.appendChild(previewLoading);
              }
              viewerType = ext;
            }

            if (!viewer) {
              previewLoading.classList.remove('hidden');
              try {
                var mod = await getModule(ext);
                await mod.init(wasmUrl(ext));
                viewer = await mod.mount(previewContainer);
              } catch (e) {
                console.error('Preview init error:', e);
                previewLoading.textContent = 'Preview unavailable';
                return;
              }
              previewLoading.classList.add('hidden');
            }

            var t0 = performance.now();
            try { viewer.load(new Uint8Array(arrayBuffer)); } catch (e) { console.error('Preview load error:', e); }
            previewTime.textContent = Math.round(performance.now() - t0) + 'ms';
            knownVersion = await fetchVersion();
          }

          async function refreshPreview() {
            try {
              var res = await fetch('/api/file');
              if (!res.ok) return;
              var buf = await res.arrayBuffer();
              if (viewer) {
                var t0 = performance.now();
                viewer.load(new Uint8Array(buf));
                previewTime.textContent = Math.round(performance.now() - t0) + 'ms';
                knownVersion = await fetchVersion();
              }
            } catch (e) { console.error('Preview refresh error:', e); }
          }

          function hidePreview() {
            previewPanel.classList.add('hidden');
            previewPanel.classList.remove('flex');
            resizeHandle.classList.add('hidden');
            if (viewer) { try { viewer.destroy(); } catch {} viewer = null; viewerType = null; }
            stopVersionPoll();
          }

          var knownVersion = '0';
          var pollInterval = null;

          async function fetchVersion() {
            try {
              var r = await fetch('/api/file-version');
              return r.ok ? await r.text() : '0';
            } catch { return '0'; }
          }

          async function checkVersion() {
            var v = await fetchVersion();
            if (v !== '0' && v !== knownVersion) { knownVersion = v; await refreshPreview(); }
          }

          function startVersionPoll() {
            if (pollInterval) return;
            pollInterval = setInterval(checkVersion, 1500);
          }

          function stopVersionPoll() {
            if (pollInterval) { clearInterval(pollInterval); pollInterval = null; }
          }

          window._showPreview = function(buf, name) { startVersionPoll(); return showPreview(buf, name); };
          window._showPreviewFromServer = async function(name) {
            try {
              var res = await fetch('/api/file');
              if (!res.ok) return;
              var buf = await res.arrayBuffer();
              startVersionPoll();
              await showPreview(buf, name);
            } catch (e) { console.error('Preview init error:', e); }
          };
          window._refreshPreview = refreshPreview;
          window._hidePreview = hidePreview;
        `}</script>
      </body>
    </html>
  );
}
