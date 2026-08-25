export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  if (req.method === "OPTIONS") return res.status(200).end();

  const SHEET_ID = "1Mw46A8j0c5-6VyOY8K7bEFPTnEGLh6nG-lL5jEXe2G0";
  const sheet = req.query.sheet || "Gerencial";
  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&sheet=${encodeURIComponent(sheet)}`;

  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Google retornou ${response.status}`);
    const csv = await response.text();
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.status(200).send(csv);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
