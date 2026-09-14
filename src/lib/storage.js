/* ------------------------------------------------------------------ */
/*  Armazenamento COMPARTILHADO — fala com /api/data, que por sua vez  */
/*  guarda os dados no Upstash Redis. Assim, qualquer pessoa que abre  */
/*  o painel gerencial ou qualquer /parceiro/... vê os mesmos dados,   */
/*  sem depender do navegador de quem fez o upload.                    */
/*                                                                      */
/*  Requer as variáveis de ambiente UPSTASH_REDIS_REST_URL e            */
/*  UPSTASH_REDIS_REST_TOKEN configuradas na Vercel — veja o README.    */
/* ------------------------------------------------------------------ */

export const storage = {
  async get(key) {
    try {
      const res = await fetch(`/api/data?key=${encodeURIComponent(key)}`);
      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        console.error("storage.get falhou:", body.error || res.status);
        return null;
      }
      const data = await res.json();
      if (data.value === null || data.value === undefined) return null;
      return { key, value: data.value };
    } catch (e) {
      console.error("storage.get falhou:", e);
      return null;
    }
  },

  async set(key, value) {
    try {
      const res = await fetch(`/api/data?key=${encodeURIComponent(key)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ value }),
      });
      if (!res.ok) {
        const body = await res.json().catch(() => ({}));
        console.error("storage.set falhou:", body.error || res.status);
        return null;
      }
      const data = await res.json();
      return { key, value: data.value };
    } catch (e) {
      console.error("storage.set falhou:", e);
      return null;
    }
  },
};
