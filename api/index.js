const URL = process.env.APPS_SCRIPT_URL;

export default async function(req, res) {
  if (!URL) {
    return res.status(500).json({
      ok: false,
      error: "APPS_SCRIPT_URL is not configured."
    });
  }

  try {
    let target = URL;

    if (req.method === "GET" && req.url.includes("?")) {
      target += req.url.slice(req.url.indexOf("?"));
    }

    const response = await fetch(target, {
      method: req.method,
      headers: {
        "Content-Type": "application/json"
      },
      body: req.method === "POST"
        ? JSON.stringify(req.body || {})
        : undefined
    });

    const text = await response.text();

    let data;
    try {
      data = JSON.parse(text);
    } catch {
      return res.status(502).json({
        ok: false,
        error: "Apps Script returned a non-JSON response.",
        upstreamStatus: response.status,
        upstreamPreview: text.slice(0, 300)
      });
    }

    return res.status(response.status).json(data);

  } catch (e) {
    return res.status(502).json({
      ok: false,
      error: String(e.message || e)
    });
  }
}
