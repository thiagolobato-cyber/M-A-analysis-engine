// Só Admin pode trocar o papel de alguém — verificado no servidor, não só
// escondendo o botão na tela (esconder botão não é segurança).

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, content-type, apikey",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

async function getCallerProfile(accessToken: string) {
  const userRes = await fetch(`${Deno.env.get("SUPABASE_URL")}/auth/v1/user`, {
    headers: {
      apikey: Deno.env.get("SUPABASE_ANON_KEY") ?? "",
      Authorization: `Bearer ${accessToken}`,
    },
  });
  if (!userRes.ok) return null;
  const user = await userRes.json();

  const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
  const profileRes = await fetch(
    `${Deno.env.get("SUPABASE_URL")}/rest/v1/profiles?id=eq.${user.id}&select=id,email,role`,
    { headers: { apikey: serviceKey, Authorization: `Bearer ${serviceKey}` } }
  );
  const profiles = await profileRes.json();
  return profiles[0] ?? null;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: corsHeaders });
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const { user_id, new_role } = await req.json().catch(() => ({}));
  if (!user_id || !["admin", "analyst"].includes(new_role)) {
    return new Response(JSON.stringify({ error: "user_id e new_role ('admin'|'analyst') são obrigatórios" }), {
      status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const accessToken = (req.headers.get("authorization") || "").replace("Bearer ", "");
  const caller = await getCallerProfile(accessToken);
  if (!caller || caller.role !== "admin") {
    return new Response(JSON.stringify({ error: "Só Admin pode alterar papéis" }), {
      status: 403, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
  const updateRes = await fetch(
    `${Deno.env.get("SUPABASE_URL")}/rest/v1/profiles?id=eq.${user_id}`,
    {
      method: "PATCH",
      headers: {
        apikey: serviceKey,
        Authorization: `Bearer ${serviceKey}`,
        "Content-Type": "application/json",
        Prefer: "return=representation",
      },
      body: JSON.stringify({ role: new_role }),
    }
  );

  if (!updateRes.ok) {
    return new Response(JSON.stringify({ error: await updateRes.text() }), {
      status: 502, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  // Registra a própria mudança de papel no log de atividade
  await fetch(`${Deno.env.get("SUPABASE_URL")}/rest/v1/activity_log`, {
    method: "POST",
    headers: { apikey: serviceKey, Authorization: `Bearer ${serviceKey}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      user_id: caller.id, user_email: caller.email, action: "change_role",
      details: { target_user_id: user_id, new_role },
    }),
  });

  return new Response(JSON.stringify({ ok: true }), {
    status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
});
