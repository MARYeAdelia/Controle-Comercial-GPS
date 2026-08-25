export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  if (req.method === "OPTIONS") return res.status(200).end();

  const SHEET_ID = "1Mw46A8j0c5-6VyOY8K7bEFPTnEGLh6nG-lL5jEXe2G0";
  const sheet = req.query.sheet || "Gerencial";

  // Tenta pelo gid também como fallback
  const urls = [
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&sheet=${encodeURIComponent(sheet)}`,
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheet)}`,
  ];

  let lastError = null;
  for (const url of urls) {
    try {
      const response = await fetch(url, {
        headers: {
          "User-Agent": "Mozilla/5.0",
          "Accept": "text/csv,text/plain,*/*",
        },
        redirect: "follow",
      });
      if (!response.ok) {
        lastError = `URL ${url} retornou ${response.status}`;
        continue;
      }
      const csv = await response.text();
      if (!csv || csv.length < 10) {
        lastError = `CSV vazio da URL ${url}`;
        continue;
      }
      res.setHeader("Content-Type", "text/plain; charset=utf-8");
      return res.status(200).send(csv);
    } catch (err) {
      lastError = err.message;
    }
  }

  res.status(500).json({ error: lastError || "Falha ao buscar planilha" });
}
