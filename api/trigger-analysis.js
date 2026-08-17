// Vercel Serverless Function — roda no servidor, nunca no navegador.
// Segura o GITHUB_TOKEN (variável de ambiente da Vercel) longe do cliente.
//
// Fluxo:
//   1. Recebe { deal_id } + o token de sessão do Supabase (Authorization header)
//   2. Confirma que o token é de um usuário autenticado de verdade
//      (não confia em nada que venha só do navegador sem checar)
//   3. Dispara o workflow_dispatch do GitHub Actions

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const { deal_id } = req.body || {};
  if (!deal_id) {
    return res.status(400).json({ error: "deal_id é obrigatório" });
  }

  const authHeader = req.headers.authorization || "";
  const accessToken = authHeader.replace("Bearer ", "");
  if (!accessToken) {
    return res.status(401).json({ error: "Sem token de autenticação" });
  }

  // Confirma que o token realmente pertence a um usuário logado no Supabase
  const userCheck = await fetch(`${process.env.SUPABASE_URL}/auth/v1/user`, {
    headers: {
      apikey: process.env.SUPABASE_ANON_KEY,
      Authorization: `Bearer ${accessToken}`,
    },
  });
  if (!userCheck.ok) {
    return res.status(401).json({ error: "Sessão inválida ou expirada" });
  }

  const dispatchRes = await fetch(
    `https://api.github.com/repos/${process.env.GITHUB_REPO}/actions/workflows/analise-ma.yml/dispatches`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${process.env.GITHUB_TOKEN}`,
        Accept: "application/vnd.github+json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ ref: "main", inputs: { deal_id } }),
    }
  );

  if (!dispatchRes.ok) {
    const text = await dispatchRes.text();
    return res.status(502).json({ error: `GitHub recusou o disparo: ${text}` });
  }

  return res.status(200).json({ ok: true });
}
