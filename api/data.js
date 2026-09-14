/* ------------------------------------------------------------------ */
/*  Função serverless da Vercel: /api/data                             */
/*  Guarda e lê dados no Upstash Redis, para que todo mundo que abre    */
/*  o link (painel gerencial ou qualquer /parceiro/...) veja os         */
/*  mesmos dados — em vez de cada navegador ter a sua própria cópia.    */
/*                                                                       */
/*  Precisa de duas variáveis de ambiente na Vercel (veja o README):    */
/*    UPSTASH_REDIS_REST_URL                                            */
/*    UPSTASH_REDIS_REST_TOKEN                                          */
/* ------------------------------------------------------------------ */

export default async function handler(req, res) {
  const { UPSTASH_REDIS_REST_URL, UPSTASH_REDIS_REST_TOKEN } = process.env;

  if (!UPSTASH_REDIS_REST_URL || !UPSTASH_REDIS_REST_TOKEN) {
    res.status(500).json({
      error: "Banco de dados não configurado. Defina UPSTASH_REDIS_REST_URL e UPSTASH_REDIS_REST_TOKEN nas variáveis de ambiente da Vercel (veja o README).",
    });
    return;
  }

  const key = (req.query.key || "").toString();
  if (!key) {
    res.status(400).json({ error: "Parâmetro 'key' é obrigatório." });
    return;
  }
  const redisKey = `logistics-dashboard:${key}`;

  try {
    if (req.method === "GET") {
      const r = await fetch(`${UPSTASH_REDIS_REST_URL}/get/${encodeURIComponent(redisKey)}`, {
        headers: { Authorization: `Bearer ${UPSTASH_REDIS_REST_TOKEN}` },
      });
      if (!r.ok) throw new Error(`Upstash respondeu ${r.status}`);
      const data = await r.json();
      res.status(200).json({ key, value: data.result ?? null });
      return;
    }

    if (req.method === "POST" || req.method === "PUT") {
      let body = req.body;
      if (typeof body === "string") {
        try { body = JSON.parse(body); } catch { body = {}; }
      }
      const value = body?.value;
      if (typeof value !== "string") {
        res.status(400).json({ error: "Corpo da requisição precisa ter um campo 'value' (string)." });
        return;
      }

      const r = await fetch(`${UPSTASH_REDIS_REST_URL}/set/${encodeURIComponent(redisKey)}`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${UPSTASH_REDIS_REST_TOKEN}`,
          "Content-Type": "text/plain",
        },
        body: value,
      });
      if (!r.ok) throw new Error(`Upstash respondeu ${r.status}`);
      const data = await r.json();
      if (data.result !== "OK") throw new Error("Upstash não confirmou a gravação.");

      res.status(200).json({ key, value });
      return;
    }

    res.status(405).json({ error: "Método não suportado." });
  } catch (e) {
    console.error("Erro em /api/data:", e);
    res.status(502).json({ error: `Falha ao falar com o banco: ${e.message}` });
  }
}
