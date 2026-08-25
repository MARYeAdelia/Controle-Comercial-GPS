export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  if (req.method === "OPTIONS") return res.status(200).end();

  const SHEET_ID = "1Mw46A8j0c5-6VyOY8K7bEFPTnEGLh6nG-lL5jEXe2G0";
  const gid = req.query.gid || "2073814116";

  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${gid}`;

  try {
    const response = await fetch(url, {
      headers: { "User-Agent": "Mozilla/5.0", "Accept": "text/csv,*/*" },
      redirect: "follow",
    });
    if (!response.ok) throw new Error(`Google retornou ${response.status}`);
    const csv = await response.text();
    if (!csv || csv.length < 10) throw new Error("CSV vazio");
    res.setHeader("Content-Type", "text/plain; charset=utf-8");
    res.status(200).send(csv);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
