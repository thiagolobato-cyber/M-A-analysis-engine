// Supabase Edge Function — roda no servidor do Supabase, nunca no navegador.
// Substitui o api/trigger-analysis.js da Vercel (removido) porque agora o
// site é hospedado no GitHub Pages (puramente estático) e essa é a única
// peça que precisa rodar em servidor, para guardar o GITHUB_TOKEN em segredo.
//
// Como o GitHub Pages e esta função vivem em domínios diferentes, é
// necessário CORS explícito (não seria preciso se tudo estivesse no mesmo
// domínio, como era o caso da versão Vercel).

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, content-type, apikey",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const { deal_id } = await req.json().catch(() => ({}));
  if (!deal_id) {
    return new Response(JSON.stringify({ error: "deal_id é obrigatório" }), {
      status: 400,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const authHeader = req.headers.get("authorization") || "";
  const accessToken = authHeader.replace("Bearer ", "");
  if (!accessToken) {
    return new Response(JSON.stringify({ error: "Sem token de autenticação" }), {
      status: 401,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  // Confirma que o token realmente pertence a um usuário logado no Supabase
  // (nunca confia em algo que só o navegador afirma)
  const userCheck = await fetch(`${Deno.env.get("SUPABASE_URL")}/auth/v1/user`, {
    headers: {
      apikey: Deno.env.get("SUPABASE_ANON_KEY") ?? "",
      Authorization: `Bearer ${accessToken}`,
    },
  });
  if (!userCheck.ok) {
    return new Response(JSON.stringify({ error: "Sessão inválida ou expirada" }), {
      status: 401,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const dispatchRes = await fetch(
    `https://api.github.com/repos/${Deno.env.get("GITHUB_REPO")}/actions/workflows/analise-ma.yml/dispatches`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${Deno.env.get("GITHUB_TOKEN")}`,
        Accept: "application/vnd.github+json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ ref: "main", inputs: { deal_id } }),
    }
  );

  if (!dispatchRes.ok) {
    const text = await dispatchRes.text();
    return new Response(JSON.stringify({ error: `GitHub recusou o disparo: ${text}` }), {
      status: 502,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  return new Response(JSON.stringify({ ok: true }), {
    status: 200,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
});
