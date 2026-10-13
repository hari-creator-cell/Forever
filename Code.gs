
const P = PropertiesService.getScriptProperties();

const SID = () => P.getProperty("SHEET_ID");
const OWNER = () => (P.getProperty("OWNER_EMAIL") || "h9302782@gmail.com").trim().toLowerCase();
const CID = () => P.getProperty("GOOGLE_CLIENT_ID");
const TZ = () => Session.getScriptTimeZone() || "Asia/Kolkata";
const EXP_FOLDER = () => P.getProperty("EXPERIENCES_FOLDER_ID") || "1i26eSke15-vMB5hvR8RBIqelw60AiZx4";

const X = {
  ID: 0, CREATOR: 1, TYPE: 2, TITLE: 3, NAME: 4, MSG: 5,
  FOLDER: 6, JSON: 7, PHOTO: 8, AUDIO: 9, CREATED: 10,
  EXPIRES: 11, VISITS: 12, MAX: 13, STATUS: 14, PRICE: 15
};

const EXP_HEADERS = [
  "experience_id", "creator_id", "type", "title", "recipient_name",
  "message", "drive_folder_id", "experience_json_file_id",
  "photo_file_id", "audio_file_id", "created_at", "expires_at",
  "visit_count", "max_visits", "status", "price"
];

const TEMPLATES = {
  "escaping-no": { type: "proposal", title: "The Escaping NO", final: true },
  "classic-birthday": { type: "birthday", title: "Classic Birthday" },
  "achievement-celebration": { type: "congratulations", title: "Achievement Celebration" },
  "simple-love-card": { type: "free-card", title: "Simple Love Card" },
  "simple-birthday-card": { type: "free-card", title: "Simple Birthday Card" },
  "simple-congrats-card": { type: "free-card", title: "Simple Congratulations Card" }
};

const LIMITS = {
  name: 40,
  msg: 600,
  fin: 200,
  photo: 800 * 1024,
  audio: 2 * 1024 * 1024
};

const EXT = {
  "image/jpeg": "jpg",
  "image/png": "png",
  "image/webp": "webp",
  "image/gif": "gif",
  "audio/mpeg": "mp3",
  "audio/mp3": "mp3",
  "audio/mp4": "m4a",
  "audio/x-m4a": "m4a",
  "audio/m4a": "m4a",
  "audio/aac": "aac",
  "audio/wav": "wav",
  "audio/x-wav": "wav",
  "audio/wave": "wav",
  "audio/ogg": "ogg",
  "audio/webm": "webm",
  "audio/flac": "flac"
};

const ERR = {
  E_AUTH: "Your session expired. Please sign in again.",
  E_FORBIDDEN: "You don't have access to that.",
  E_INVALID: "Some details are missing or invalid. Please check the form.",
  E_TOO_LARGE: "One of your files is too large.",
  E_DRIVE: "We couldn't save your files right now. Please try again.",
  E_SHEET: "We couldn't save your experience right now. Please try again.",
  E_NOT_FOUND: "This experience couldn't be found.",
  E_UNAVAILABLE: "This experience is no longer available.",
  E_EXPIRED: "This experience has expired.",
  E_LIMIT: "This experience has reached its visit limit.",
  E_SUSPENDED: "This account has been suspended. Contact the site owner.",
  E_SERVER: "Something went wrong on our side. Please try again."
};

function die(code) {
  throw new Error(code);
}

function out(x) {
  return ContentService
    .createTextOutput(JSON.stringify(x))
    .setMimeType(ContentService.MimeType.JSON);
}

function errOut(x) {
  const m = String((x && x.message) || x);
  const code = ERR[m] ? m : "E_SERVER";

  if (code === "E_SERVER") {
    Logger.log("Forever error: " + m + (x && x.stack ? "\n" + x.stack : ""));
  }

  return out({ ok: false, success: false, code: code, error: ERR[code] });
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const a = p.action || "";

    if (a === "health") {
      return out({
        ok: true,
        service: "Forever",
        status: "UP",
        timestamp: new Date().toISOString()
      });
    }

    if (a === "experience") return out(openExp(p.id));
    if (a === "getMedia") return out(getMedia(p.id, p.kind));
    if (p.p) return out(openExp(p.p));

    return out({ ok: true, service: "Forever" });
  } catch (x) {
    return errOut(x);
  }
}

function doPost(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    if (b.action === "login") return out(login(b.idToken));
    if (b.action === "createExperience") return out(create(b));
    if (b.action === "adminStats") return out(stats(b.idToken));

    if (b.action === "adminData") {
      return out(adminData(b.idToken, b.section));
    }

    if (b.action === "adminExperienceStatus") {
      return out(adminExperienceStatus(b.idToken, b.id, b.status));
    }

    if (b.action === "adminUserStatus") {
      return out(adminUserStatus(b.idToken, b.email, b.status));
    }

    die("E_INVALID");
  } catch (x) {
    return errOut(x);
  }
}

// ---------- Google identity verification ----------

function verify(t) {
  if (!t) die("E_AUTH");

  const r = UrlFetchApp.fetch(
    "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(t),
    { muteHttpExceptions: true }
  );

  if (r.getResponseCode() !== 200) die("E_AUTH");

  const g = JSON.parse(r.getContentText());

  if (g.aud !== CID() || !g.email) die("E_AUTH");
  if (String(g.email_verified).toLowerCase() !== "true") die("E_AUTH");

  return {
    email: String(g.email).trim().toLowerCase(),
    name: g.name || g.email,
    sub: g.sub || ""
  };
}

function sh(n) {
  const ss = SpreadsheetApp.openById(SID());
  let s = ss.getSheetByName(n);
  if (!s) s = ss.insertSheet(n);
  return s;
}

function rows(n) {
  return sh(n).getDataRange().getValues();
}

function now() {
  return new Date();
}

function clean(s, max) {
  return String(s == null ? "" : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .trim()
    .slice(0, max);
}

function userRowFor_(email) {
  const r = rows("Users");

  for (let i = 1; i < r.length; i++) {
    if (String(r[i][2] || "").trim().toLowerCase() === email) {
      return { row: i + 1, values: r[i] };
    }
  }

  return null;
}

function requireActiveUser_(email) {
  if (email === OWNER()) return;

  const found = userRowFor_(email);

  if (found && String(found.values[9] || "active").toLowerCase() === "suspended") {
    die("E_SUSPENDED");
  }
}

function userIdFor(email) {
  const found = userRowFor_(email);
  return found ? String(found.values[0]) : email;
}

// ---------- login ----------

function login(t) {
  const g = verify(t);
  const s = sh("Users");
  const found = userRowFor_(g.email);
  const n = now();

  if (found && String(found.values[9] || "active").toLowerCase() === "suspended" && g.email !== OWNER()) {
    die("E_SUSPENDED");
  }

  if (!found) {
    s.appendRow([
      "usr_" + Utilities.getUuid().slice(0, 8),
      g.sub,
      g.email,
      g.name,
      "",
      n,
      n,
      1,
      g.email === OWNER() ? "owner" : "user",
      "active"
    ]);
  } else {
    s.getRange(found.row, 7).setValue(n);
    s.getRange(found.row, 8).setValue(Number(found.values[7] || 0) + 1);

    // Keep owner role consistent with the server-side owner email.
    if (g.email === OWNER()) {
      s.getRange(found.row, 9).setValue("owner");
    }
  }

  sh("Events").appendRow([
    "evt_" + Utilities.getUuid().slice(0, 8),
    g.email,
    "",
    "login",
    n,
    "",
    g.email
  ]);

  return {
    ok: true,
    success: true,
    user: {
      email: g.email,
      name: g.name,
      role: g.email === OWNER() ? "owner" : "user"
    }
  };
}

// ---------- creation helpers ----------

function ensureExpHeader() {
  const s = sh("Experiences");

  if (s.getLastRow() === 0) {
    s.appendRow(EXP_HEADERS);
    return s;
  }

  if (
    String(s.getRange(1, 1).getValue()) !== "experience_id" ||
    s.getLastColumn() < EXP_HEADERS.length
  ) {
    Logger.log("Experiences header mismatch");
    die("E_SHEET");
  }

  return s;
}

function newExpId(existing) {
  const A = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";

  for (let k = 0; k < 20; k++) {
    const hex = Utilities.getUuid().replace(/-/g, "");
    let id = "EXP_";

    for (let i = 0; i < 6; i++) {
      id += A[parseInt(hex.substr(i * 2, 2), 16) % 32];
    }

    if (!existing[id]) return id;
  }

  die("E_SERVER");
}

function saveFile_(folder, f, kind) {
  if (!f || !f.data) return "";

  const ok = kind === "photo"
    ? /^image\//.test(f.type || "")
    : /^audio\//.test(f.type || "");

  const ext = EXT[f.type];

  if (!ok || !ext) die("E_INVALID");

  let bytes;

  try {
    bytes = Utilities.base64Decode(f.data);
  } catch (e) {
    die("E_INVALID");
  }

  if (bytes.length > (kind === "photo" ? LIMITS.photo : LIMITS.audio)) {
    die("E_TOO_LARGE");
  }

  return folder.createFile(
    Utilities.newBlob(bytes, f.type, kind + "." + ext)
  ).getId();
}

// ---------- create experience ----------

function create(b) {
  const g = verify(b.idToken);
  requireActiveUser_(g.email);

  const e = b.experience || {};
  const tpl = TEMPLATES[e.template];

  if (!tpl) die("E_INVALID");

  const name = clean(e.recipientName, LIMITS.name);
  const msg = clean(e.message, LIMITS.msg);
  const fin = tpl.final
    ? (clean(e.finalMessage, LIMITS.fin) || "You said YES! ❤️")
    : "";

  if (!name || !msg) die("E_INVALID");

  const files = b.files || {};
  const lock = LockService.getScriptLock();

  lock.waitLock(20000);

  let folder = null;

  try {
    const sheet = ensureExpHeader();
    const existing = {};

    sheet.getDataRange().getValues().forEach(r => {
      existing[String(r[X.ID])] = 1;
    });

    const id = newExpId(existing);
    const created = now();

    let photoId = "";
    let audioId = "";
    let jsonId = "";

    try {
      folder = DriveApp.getFolderById(EXP_FOLDER()).createFolder(id);

      photoId = saveFile_(folder, files.photo, "photo");
      audioId = saveFile_(folder, files.audio, "audio");

      const cfg = {
        experience_id: id,
        type: tpl.type,
        template: e.template,
        recipient_name: name,
        message: msg,
        final_message: fin,
        photo_file_id: photoId,
        audio_file_id: audioId,
        created_at: created.toISOString(),
        status: "active"
      };

      jsonId = folder.createFile(
        "experience.json",
        JSON.stringify(cfg, null, 2),
        MimeType.PLAIN_TEXT
      ).getId();

    } catch (x) {
      if (folder) {
        try { folder.setTrashed(true); } catch (_) {}
      }

      const m = String(x && x.message);
      if (ERR[m]) throw x;

      Logger.log("Drive failure: " + m);
      die("E_DRIVE");
    }

    try {
      sheet.appendRow([
        id,
        userIdFor(g.email),
        tpl.type,
        tpl.title,
        name,
        msg,
        folder.getId(),
        jsonId,
        photoId,
        audioId,
        created,
        "",
        0,
        0,
        "active",
        0
      ]);

      sh("Events").appendRow([
        "evt_" + Utilities.getUuid().slice(0, 8),
        g.email,
        id,
        "experience_created",
        created,
        "",
        tpl.title
      ]);

    } catch (x) {
      Logger.log("Sheet failure: " + String(x && x.message));

      try { folder.setTrashed(true); } catch (_) {}

      die("E_SHEET");
    }

    return {
      ok: true,
      success: true,
      id: id,
      url: "/p/" + id
    };

  } finally {
    lock.releaseLock();
  }
}

// ---------- experience lookup and recipient access ----------

function validId(id) {
  return /^[A-Za-z0-9_]{4,24}$/.test(String(id || ""));
}

function findExp(id) {
  const s = sh("Experiences");
  const r = s.getDataRange().getValues();

  for (let i = 1; i < r.length; i++) {
    if (String(r[i][X.ID]) === String(id)) {
      return { row: i + 1, v: r[i], sheet: s };
    }
  }

  return null;
}

function checkLive(v) {
  if (String(v[X.STATUS]) !== "active") die("E_UNAVAILABLE");

  const exp = v[X.EXPIRES];

  if (exp && new Date(exp) < now()) die("E_EXPIRED");
}

function readCfg(v) {
  try {
    return JSON.parse(
      DriveApp.getFileById(String(v[X.JSON])).getBlob().getDataAsString()
    ) || {};
  } catch (e) {
    return {};
  }
}

function openExp(id) {
  if (!validId(id)) die("E_NOT_FOUND");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const f = findExp(id);

    if (!f) die("E_NOT_FOUND");

    const v = f.v;
    checkLive(v);

    const visits = Number(v[X.VISITS] || 0);
    const max = Number(v[X.MAX] || 0);

    if (max > 0 && visits >= max) die("E_LIMIT");

    f.sheet.getRange(f.row, X.VISITS + 1).setValue(visits + 1);

    sh("Events").appendRow([
      "evt_" + Utilities.getUuid().slice(0, 8),
      "",
      String(v[X.ID]),
      "experience_opened",
      now(),
      "",
      ""
    ]);

    const cfg = readCfg(v);

    return {
      ok: true,
      success: true,
      experience: {
        id: String(v[X.ID]),
        template: cfg.template || (
          String(v[X.TYPE]) === "proposal"
            ? "escaping-no"
            : "simple-love-card"
        ),
        recipientName: String(v[X.NAME]),
        message: String(v[X.MSG]),
        finalMessage: String(cfg.final_message || ""),
        hasPhoto: !!v[X.PHOTO],
        hasAudio: !!v[X.AUDIO]
      }
    };

  } finally {
    lock.releaseLock();
  }
}

function getMedia(id, kind) {
  if (!validId(id) || (kind !== "photo" && kind !== "audio")) {
    die("E_NOT_FOUND");
  }

  const f = findExp(id);

  if (!f) die("E_NOT_FOUND");

  checkLive(f.v);

  const fileId = String(
    f.v[kind === "photo" ? X.PHOTO : X.AUDIO] || ""
  );

  if (!fileId) die("E_NOT_FOUND");

  const blob = DriveApp.getFileById(fileId).getBlob();

  return {
    ok: true,
    success: true,
    dataUri: "data:" + blob.getContentType() +
      ";base64," + Utilities.base64Encode(blob.getBytes())
  };
}

// ---------- statistics ----------

function fmt(x) {
  if (!x) return "";

  const d = x instanceof Date ? x : new Date(x);

  if (isNaN(d.getTime())) return "";

  return Utilities.formatDate(d, TZ(), "yyyy-MM-dd");
}

function stats(t) {
  ownerCheck_(t);

  const u = rows("Users");
  const e = rows("Experiences");
  const ev = rows("Events");
  const o = rows("Orders");
  const today = Utilities.formatDate(now(), TZ(), "yyyy-MM-dd");

  let nu = 0;
  let li = 0;
  let vi = 0;
  let views = 0;
  let rv = 0;
  let tr = 0;

  u.slice(1).forEach(x => {
    if (fmt(x[5]) === today) nu++;
  });

  ev.slice(1).forEach(x => {
    if (fmt(x[4]) === today) {
      if (x[3] === "login") li++;

      if (x[3] === "experience_opened") {
        views++;
        vi++;
      }
    }
  });

  const map = {};

  o.slice(1).forEach(x => {
    if (String(x[4]).toLowerCase() === "paid") {
      rv += Number(x[3] || 0);

      if (fmt(x[5]) === today) tr += Number(x[3] || 0);

      const id = String(x[2]);
      map[id] = (map[id] || 0) + 1;
    }
  });

  const top = e.slice(1).map(x => ({
    id: String(x[0]),
    title: String(x[3]),
    orders: map[String(x[0])] || 0
  }))
    .filter(x => x.orders > 0)
    .sort((a, b) => b.orders - a.orders)
    .slice(0, 10);

  return {
    ok: true,
    success: true,
    stats: {
      overall: {
        users: Math.max(0, u.length - 1),
        visitors: ev.slice(1).filter(x => x[3] === "experience_opened").length,
        experiences: Math.max(0, e.length - 1),
        revenue: rv
      },
      today: {
        newUsers: nu,
        logins: li,
        visitors: vi,
        views: views,
        revenue: tr
      },
      topExperiences: top
    }
  };
}

// ---------- owner-only dashboard APIs ----------

function ownerCheck_(token) {
  const g = verify(token);

  if (g.email !== OWNER()) die("E_FORBIDDEN");

  return g;
}

function dateLabel_(v) {
  if (!v) return "";

  const d = v instanceof Date ? v : new Date(v);

  if (isNaN(d.getTime())) return "";

  return Utilities.formatDate(d, TZ(), "yyyy-MM-dd HH:mm");
}

function adminData(token, section) {
  ownerCheck_(token);

  section = String(section || "experiences");

  if (section === "experiences") {
    const r = rows("Experiences");

    return {
      ok: true,
      success: true,
      section: section,
      items: r.slice(1)
        .filter(x => x[X.ID])
        .map(x => ({
          id: String(x[X.ID]),
          creator: String(x[X.CREATOR] || ""),
          type: String(x[X.TYPE] || ""),
          title: String(x[X.TITLE] || ""),
          recipient: String(x[X.NAME] || ""),
          created: dateLabel_(x[X.CREATED]),
          visits: Number(x[X.VISITS] || 0),
          status: String(x[X.STATUS] || "active"),
          price: Number(x[X.PRICE] || 0)
        }))
        .reverse()
    };
  }

  if (section === "users") {
    const r = rows("Users");

    return {
      ok: true,
      success: true,
      section: section,
      items: r.slice(1)
        .filter(x => x[0])
        .map(x => ({
          id: String(x[0]),
          email: String(x[2] || ""),
          name: String(x[3] || ""),
          created: dateLabel_(x[5]),
          lastLogin: dateLabel_(x[6]),
          logins: Number(x[7] || 0),
          role: String(x[8] || "user"),
          status: String(x[9] || "active")
        }))
        .reverse()
    };
  }

  if (section === "orders") {
    const r = rows("Orders");

    return {
      ok: true,
      success: true,
      section: section,
      items: r.slice(1)
        .filter(x => x[0])
        .map(x => ({
          id: String(x[0]),
          user: String(x[1] || ""),
          experience: String(x[2] || ""),
          amount: Number(x[3] || 0),
          payment: String(x[4] || "pending"),
          created: dateLabel_(x[5])
        }))
        .reverse()
    };
  }

  if (section === "settings") {
    return {
      ok: true,
      success: true,
      section: section,
      owner: OWNER(),
      service: "Forever",
      storage: "Google Drive",
      database: "Google Sheets",
      payments: "Not configured",
      spreadsheetConfigured: !!SID(),
      googleClientConfigured: !!CID(),
      experiencesFolderConfigured: !!EXP_FOLDER()
    };
  }

  die("E_INVALID");
}

function adminExperienceStatus(token, id, status) {
  ownerCheck_(token);

  id = String(id || "");
  status = String(status || "");

  if (
    !validId(id) ||
    ["active", "inactive", "removed"].indexOf(status) < 0
  ) {
    die("E_INVALID");
  }

  const f = findExp(id);

  if (!f) die("E_NOT_FOUND");

  f.sheet.getRange(f.row, X.STATUS + 1).setValue(status);

  sh("Events").appendRow([
    "evt_" + Utilities.getUuid().slice(0, 8),
    OWNER(),
    id,
    "admin_experience_" + status,
    now(),
    "",
    "Owner dashboard"
  ]);

  return {
    ok: true,
    success: true,
    id: id,
    status: status
  };
}

function adminUserStatus(token, email, status) {
  ownerCheck_(token);

  email = String(email || "").trim().toLowerCase();
  status = String(status || "").trim().toLowerCase();

  if (
    !email ||
    ["active", "suspended"].indexOf(status) < 0
  ) {
    die("E_INVALID");
  }

  if (email === OWNER()) die("E_FORBIDDEN");

  const s = sh("Users");
  const r = s.getDataRange().getValues();

  for (let i = 1; i < r.length; i++) {
    if (String(r[i][2] || "").trim().toLowerCase() === email) {
      s.getRange(i + 1, 10).setValue(status);

      sh("Events").appendRow([
        "evt_" + Utilities.getUuid().slice(0, 8),
        OWNER(),
        "",
        "admin_user_" + status,
        now(),
        "",
        email
      ]);

      return {
        ok: true,
        success: true,
        email: email,
        status: status
      };
    }
  }

  die("E_NOT_FOUND");
}

// ---------- one-time sheet setup ----------

function setup() {
  const ss = SpreadsheetApp.openById(SID());

  const m = {
    Users: [
      "user_id", "google_id", "email", "name", "photo_url",
      "created_at", "last_login", "login_count", "role", "status"
    ],
    Experiences: EXP_HEADERS,
    Events: [
      "event_id", "user_id", "experience_id", "event_type",
      "timestamp", "session_id", "metadata"
    ],
    Orders: [
      "order_id", "user_id", "experience_id",
      "amount", "payment_status", "created_at"
    ],
    DailyStats: [
      "date", "new_users", "total_logins", "unique_visitors",
      "experience_views", "experiences_created", "orders", "revenue"
    ]
  };

  Object.keys(m).forEach(n => {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);

    if (s.getLastRow() === 0) {
      s.appendRow(m[n]);
    }
  });
}
