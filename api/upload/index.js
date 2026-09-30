module.exports = async function (context, req) {
  const token = req.headers["x-graph-token"];
  const parentId = req.query.parentId;
  const name = req.query.name;
  const copyParentId = req.query.copyParentId;
  const copyName = req.query.copyName;
  const reply = (status, error) => { context.res = { status, headers: { "Content-Type": "application/json" }, body: JSON.stringify(error ? { error } : { uploaded: true }) }; };
  if (!token || !parentId || !name) return reply(400, "Missing upload information.");

  const contentLength = Number(req.headers["content-length"] || 0);
  if (contentLength > 250 * 1024 * 1024) return reply(413, "This file is larger than 250 MB.");

  try {
    let bytes;
    if (Buffer.isBuffer(req.body)) bytes = req.body;
    else if (Buffer.isBuffer(req.rawBody)) bytes = req.rawBody;
    else if (typeof req.rawBody === "string") bytes = Buffer.from(req.rawBody, "binary");
    else if (req.body instanceof ArrayBuffer) bytes = Buffer.from(req.body);
    else bytes = Buffer.from([]);
    if (!bytes.length) return reply(400, "The uploaded photo was empty.");

    const targets = [{ parentId, name }];
    if (copyParentId && copyName) targets.push({ parentId: copyParentId, name: copyName });
    const responses = await Promise.all(targets.map((target) => fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(target.parentId)}:/${encodeURIComponent(target.name)}:/content`, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": req.headers["content-type"] || "application/octet-stream",
        },
        body: bytes,
      })));
    for (const response of responses) {
      const text = await response.text();
      if (!response.ok) {
        let message = `OneDrive returned ${response.status}`;
        try { message = JSON.parse(text)?.error?.message || message; } catch { /* Keep the HTTP status. */ }
        return reply(response.status, message);
      }
    }
    return reply(200);
  } catch (error) {
    return reply(502, error instanceof Error ? error.message : "Azure could not forward the photo to OneDrive.");
  }
};
