/* Forever experience engine – used by the live preview (index.html) AND the recipient page (player.html),
   so the preview is the real experience. */
(function (root) {
  "use strict";

  var esc = function (s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  };

  var TEMPLATES = {
    "escaping-no": { kind: "game", emoji: "💗", cta: "Open it 💗" },
    "classic-birthday": { kind: "card", emoji: "🎂", head: function (n) { return "Happy Birthday, " + n + "!"; }, confetti: ["🎈", "🎉", "✨", "🎂"], cta: "Open your surprise 🎁" },
    "achievement-celebration": { kind: "card", emoji: "🏆", head: function (n) { return "Congratulations, " + n + "!"; }, confetti: ["🎉", "🏆", "✨", "👏"], cta: "See the news 🎉" },
    "simple-love-card": { kind: "card", emoji: "💌", head: function (n) { return "For " + n + " 💗"; }, cta: "Open your card 💌" },
    "simple-birthday-card": { kind: "card", emoji: "🎁", head: function (n) { return "Happy Birthday, " + n + "!"; }, cta: "Open your card 🎁" },
    "simple-congrats-card": { kind: "card", emoji: "🎉", head: function (n) { return "Congratulations, " + n + "!"; }, cta: "Open your card 🎉" }
  };

  var TAUNTS = ["Be honest… 😌", "Nope, too slow 😏", "Missed me!", "Are you sure? 🥺", "Try again 😜", "Nice try!", "Too quick for you 💨", "Really? 🙈", "Almost! …not.", "I'm shy 🫣", "Ooh, so close", "Keep going 😆", "You can't catch me", "Still no?", "Think about it 💭", "Last chances…", "Seriously? 🥹", "Okay that's brave", "Just say yes 💗", "One more…?", "Final answer?"];
  var MAX_NO = 21;

  var CSS = ".fx{color:#fff;font-family:system-ui,-apple-system,'Segoe UI',sans-serif;text-align:center;width:100%;box-sizing:border-box}" +
    ".fx *{box-sizing:border-box}" +
    ".fx-card{position:relative;overflow:hidden;background:#141625ee;border:1px solid #30354c;border-radius:28px;padding:28px 20px;box-shadow:0 20px 80px #0008}" +
    ".fx-emoji{font-size:60px;line-height:1.1}" +
    ".fx h1{font-size:28px;line-height:1.15;margin:10px 0 8px;overflow-wrap:anywhere}" +
    ".fx-msg{white-space:pre-wrap;color:#d9dce7;font-size:18px;line-height:1.55;overflow-wrap:anywhere;margin:8px 0}" +
    ".fx-final{white-space:pre-wrap;color:#ffd1e3;font-weight:700;font-size:19px;margin:14px 0 4px;overflow-wrap:anywhere}" +
    ".fx-sub{color:#aeb4c7;font-size:14px;margin:8px 0}" +
    ".fx-photo{display:block;max-width:100%;max-height:320px;border-radius:18px;margin:16px auto;object-fit:cover}" +
    ".fx-btn{border:0;border-radius:999px;padding:14px 26px;font-weight:800;font-size:16px;color:#fff;cursor:pointer;font-family:inherit;min-height:48px}" +
    ".fx-btn:focus-visible{outline:3px solid #ffd1e3;outline-offset:2px}" +
    ".fx-yes{background:linear-gradient(135deg,#ff5b9d,#ff8a77)}" +
    ".fx-no{background:#292e42;border:1px solid #3a4058}" +
    ".fx-snd{background:#24283a;border:1px solid #3a4058;margin-top:14px;padding:10px 18px;font-size:14px;min-height:40px}" +
    ".fx-stage{position:relative;height:300px;max-width:440px;margin:12px auto 0}" +
    ".fx-stage .fx-yes{position:absolute;left:0;right:0;bottom:10px;margin:0 auto;width:max-content;z-index:2}" +
    ".fx-stage .fx-no{position:absolute;left:0;top:0;touch-action:none;user-select:none;-webkit-user-select:none;z-index:1}" +
    ".fx-pop{animation:fx-pop .22s ease-out}" +
    "@keyframes fx-pop{from{opacity:.15;transform:scale(.6)}to{opacity:1;transform:scale(1)}}" +
    ".fx-finale{animation:fx-in .5s ease}" +
    "@keyframes fx-in{from{transform:scale(.85);opacity:0}to{transform:scale(1);opacity:1}}" +
    ".fx-conf{position:absolute;inset:0;pointer-events:none;overflow:hidden}" +
    ".fx-conf span{position:absolute;top:-30px;font-size:24px;animation:fx-fall linear forwards}" +
    "@keyframes fx-fall{to{transform:translateY(560px) rotate(300deg);opacity:0}}" +
    "@media (prefers-reduced-motion:reduce){.fx-pop,.fx-finale,.fx-conf span{animation:none}}";

  function ensureCss() {
    if (typeof document === "undefined" || document.getElementById("fx-css")) return;
    var s = document.createElement("style");
    s.id = "fx-css"; s.textContent = CSS; document.head.appendChild(s);
  }

  /* Pure placement logic (unit-tested). Returns [x,y] for NO inside the stage such that NO is at least
     `M` px away from YES on at least one side (so the two can never overlap or be mis-tapped),
     preferring spots far from `from`. Returns null if the stage has no legal spot. */
  function pickPos(sw, sh, nw, nh, yes, from, M, rnd) {
    function ok(x, y) {
      return x >= 0 && y >= 0 && x + nw <= sw && y + nh <= sh &&
        (x + nw <= yes.x - M || x >= yes.x + yes.w + M || y + nh <= yes.y - M || y >= yes.y + yes.h + M);
    }
    var c = [];
    for (var y = 0; y <= sh - nh; y += 10) for (var x = 0; x <= sw - nw; x += 10) if (ok(x, y)) c.push([x, y]);
    if (!c.length) return null;
    c.sort(function (a, b) { return Math.hypot(b[0] - from.x, b[1] - from.y) - Math.hypot(a[0] - from.x, a[1] - from.y); });
    var pool = c.slice(0, Math.max(1, Math.ceil(c.length * 0.3)));
    return pool[Math.floor(rnd() * pool.length)];
  }

  function render(el, cfg, opts) {
    opts = opts || {};
    ensureCss();
    if (el.__fx) el.__fx.destroy();

    var t = TEMPLATES[cfg.template] || TEMPLATES["simple-love-card"];
    var name = (cfg.recipientName || "").trim() || "you";
    var message = (cfg.message || "").trim() || (opts.preview ? "Your message will appear here…" : "Something beautiful was made just for you.");
    var finalMsg = (cfg.finalMessage || "").trim() || "You said YES! ❤️";
    var audio = null, muted = false, ro = null, timers = [], destroyed = false, onWinResize = null;

    el.classList.add("fx");

    function later(fn, ms) { var id = setTimeout(fn, ms); timers.push(id); return id; }

    function soundHtml() { return cfg.audioUrl ? '<button class="fx-btn fx-snd" data-snd type="button">' + (opts.preview ? "▶ Play sound" : "🔊 Sound on") + "</button>" : ""; }

    function wireSound() {
      var b = el.querySelector("[data-snd]");
      if (!b || !audio) return;
      b.onclick = function () {
        if (opts.preview) {
          if (audio.paused) { audio.play().catch(function () { }); b.textContent = "⏸ Pause sound"; }
          else { audio.pause(); b.textContent = "▶ Play sound"; }
        } else {
          muted = !muted; audio.muted = muted;
          b.textContent = muted ? "🔇 Sound off" : "🔊 Sound on";
          if (!muted && audio.paused) audio.play().catch(function () { });
        }
      };
      if (!opts.preview && muted) b.textContent = "🔇 Sound off";
    }

    function confetti(host, list, n) {
      var box = document.createElement("div"); box.className = "fx-conf";
      for (var i = 0; i < n; i++) {
        var s = document.createElement("span");
        s.textContent = list[i % list.length];
        s.style.left = Math.random() * 100 + "%";
        s.style.animationDuration = 2.4 + Math.random() * 2.6 + "s";
        s.style.animationDelay = Math.random() * 1.2 + "s";
        box.appendChild(s);
      }
      host.appendChild(box);
    }

    function photoHtml() { return cfg.photoUrl ? '<img class="fx-photo" src="' + esc(cfg.photoUrl) + '" alt="A special memory">' : ""; }

    /* ---- intro (also gives the browser the tap it needs to start audio) ---- */
    function intro() {
      el.innerHTML = '<div class="fx-card"><div class="fx-emoji">✉️</div><h1>Something for ' + esc(name) + '</h1><p class="fx-sub">Tap to open</p><button class="fx-btn fx-yes" data-open type="button">' + esc(t.cta) + "</button></div>";
      el.querySelector("[data-open]").onclick = function () {
        if (audio) audio.play().catch(function () { });
        main();
      };
    }

    function main() { if (t.kind === "game") game(); else card(); }

    /* ---- simple card / birthday / congrats ---- */
    function card() {
      el.innerHTML = '<div class="fx-card"><div class="fx-emoji">' + t.emoji + "</div><h1>" + esc(t.head(name)) + '</h1><p class="fx-msg">' + esc(message) + "</p>" + photoHtml() + soundHtml() + '<p class="fx-sub" style="margin-top:18px">Made with Forever 💗</p></div>';
      if (t.confetti && !opts.preview) confetti(el.firstChild, t.confetti, 26);
      else if (t.confetti) confetti(el.firstChild, t.confetti, 8);
      wireSound();
    }

    /* ---- The Escaping NO ---- */
    function game() {
      var attempts = 0, placed = false, lastFlee = 0, done = false;
      el.innerHTML = '<div class="fx-card"><div class="fx-emoji">' + t.emoji + "</div><h1>Hey " + esc(name) + '…</h1><p class="fx-msg">' + esc(message) + "</p>" + photoHtml() +
        '<div id="fx-game"><div class="fx-stage"><button class="fx-btn fx-no" type="button">NO</button><button class="fx-btn fx-yes" type="button">YES ❤️</button></div>' +
        '<p class="fx-sub" data-taunt>' + esc(TAUNTS[0]) + '</p><p class="fx-sub" data-count style="opacity:.6">NO attempts: 0 / ' + MAX_NO + "</p></div>" + soundHtml() + "</div>";
      wireSound();

      var stage = el.querySelector(".fx-stage"), no = el.querySelector(".fx-no"), yes = el.querySelector(".fx-yes");
      var taunt = el.querySelector("[data-taunt]"), count = el.querySelector("[data-count]");
      var MARGIN = 24;

      function geo() {
        var s = stage.getBoundingClientRect(), y = yes.getBoundingClientRect(), n = no.getBoundingClientRect();
        return { sw: s.width, sh: s.height, nw: n.width, nh: n.height, yes: { x: y.left - s.left, y: y.top - s.top, w: y.width, h: y.height }, cur: { x: n.left - s.left, y: n.top - s.top } };
      }
      function place(x, y, animate) {
        no.style.left = Math.round(x) + "px"; no.style.top = Math.round(y) + "px";
        if (animate) { no.classList.remove("fx-pop"); void no.offsetWidth; no.classList.add("fx-pop"); }
      }
      function legal(g, x, y) {
        return x >= -1 && y >= -1 && x + g.nw <= g.sw + 1 && y + g.nh <= g.sh + 1 &&
          (x + g.nw <= g.yes.x - MARGIN + 1 || x >= g.yes.x + g.yes.w + MARGIN - 1 || y + g.nh <= g.yes.y - MARGIN + 1 || y >= g.yes.y + g.yes.h + MARGIN - 1);
      }
      function safeSpot(g, from) {
        var p = pickPos(g.sw, g.sh, g.nw, g.nh, g.yes, from, MARGIN, Math.random);
        if (p) return p;
        // Degenerate (tiny) stage: use the corner farthest from YES.
        var best = null, bd = -1, cx = g.yes.x + g.yes.w / 2, cy = g.yes.y + g.yes.h / 2;
        [[0, 0], [g.sw - g.nw, 0], [0, g.sh - g.nh], [g.sw - g.nw, g.sh - g.nh]].forEach(function (c) {
          var d = Math.hypot(c[0] + g.nw / 2 - cx, c[1] + g.nh / 2 - cy); if (d > bd) { bd = d; best = c; }
        });
        return best;
      }
      // Initial placement + re-validation whenever the stage changes size (e.g. mobile preview opens, rotation).
      function settle() {
        if (destroyed || done) return;
        var g = geo(); if (g.sw < 10 || g.nw < 5) return;
        if (!placed) {
          var x = (g.sw - g.nw) / 2, y = 30;
          if (!legal(g, x, y)) { var p = safeSpot(g, { x: x, y: y }); x = p[0]; y = p[1]; }
          place(x, y, false); placed = true;
        } else if (!legal(g, g.cur.x, g.cur.y)) {
          var q = safeSpot(g, g.cur); place(q[0], q[1], false);
        }
      }
      if (typeof ResizeObserver !== "undefined") { ro = new ResizeObserver(settle); ro.observe(stage); }
      else { onWinResize = settle; window.addEventListener("resize", onWinResize); }
      later(settle, 0);
      if (typeof requestAnimationFrame !== "undefined") requestAnimationFrame(settle);

      function flee() {
        if (done) return;
        var t0 = Date.now(); if (t0 - lastFlee < 200) return; lastFlee = t0;
        attempts++;
        count.textContent = "NO attempts: " + attempts + " / " + MAX_NO;
        if (attempts >= MAX_NO) { finish(false); return; }
        taunt.textContent = TAUNTS[Math.min(attempts, TAUNTS.length - 1)];
        var g = geo();
        // Teleport (no sliding transition) so NO can never visibly pass over YES.
        var p = safeSpot(g, g.cur); place(p[0], p[1], true);
      }

      no.addEventListener("pointerdown", function (ev) { ev.preventDefault(); ev.stopPropagation(); flee(); });
      no.addEventListener("pointerenter", function (ev) { if (ev.pointerType === "mouse") flee(); });
      no.addEventListener("click", function (ev) { ev.preventDefault(); ev.stopPropagation(); if (ev.detail === 0) flee(); }); // keyboard Enter/Space
      yes.addEventListener("click", function () { finish(true); });

      function finish(fromYes) {
        if (done) return; done = true;
        var host = el.querySelector("#fx-game");
        host.innerHTML = '<div class="fx-finale"><div class="fx-emoji">✨💗✨</div><h1>' + (fromYes ? "YES! 💗" : "After " + MAX_NO + " NOs… 💗") + '</h1><p class="fx-final">' + esc(finalMsg) + '</p><p class="fx-sub">' + (fromYes ? "Forever starts here." : "Okay okay, I think that's a yes 😌") + "</p></div>";
        confetti(el.firstChild, ["💗", "✨", "🎉", "💖", "🌸"], opts.preview ? 14 : 36);
      }
    }

    /* ---- go ---- */
    if (cfg.audioUrl && typeof Audio !== "undefined") {
      audio = new Audio(cfg.audioUrl); audio.loop = true; audio.volume = 0.5;
    }
    if (opts.preview || !cfg.audioUrl) main(); else intro();

    el.__fx = {
      destroy: function () {
        destroyed = true;
        timers.forEach(clearTimeout);
        if (ro) ro.disconnect();
        if (onWinResize) window.removeEventListener("resize", onWinResize);
        if (audio) { audio.pause(); audio = null; }
        el.__fx = null; el.innerHTML = "";
      }
    };
    return el.__fx;
  }

  var api = { render: render, TEMPLATES: TEMPLATES, MAX_NO: MAX_NO, _pickPos: pickPos, esc: esc };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.ForeverExp = api;
})(typeof window !== "undefined" ? window : globalThis);
