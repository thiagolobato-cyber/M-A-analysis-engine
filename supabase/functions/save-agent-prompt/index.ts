// Só Admin pode editar prompt de agente. Nunca sobrescreve — cada edição
// vira uma versão nova (mesmo princípio da seção 5 da arquitetura:
// rastreabilidade total, versionamento imutável).

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, content-type, apikey",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function bumpVersion(current: string): string {
  const match = current.match(/^v(\d+)\.(\d+)$/);
  if (!match) return "v1.1";
  const [, major, minor] = match;
  return `v${major}.${parseInt(minor) + 1}`;
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: corsHeaders });
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const { agent_name, system_prompt, model } = await req.json().catch(() => ({}));
  if (!agent_name || !system_prompt) {
    return new Response(JSON.stringify({ error: "agent_name e system_prompt são obrigatórios" }), {
      status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const supabaseUrl = Deno.env.get("SUPABASE_URL") ?? "";
  const serviceKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY") ?? "";
  const anonKey = Deno.env.get("SUPABASE_ANON_KEY") ?? "";

  const accessToken = (req.headers.get("authorization") || "").replace("Bearer ", "");
  const userRes = await fetch(`${supabaseUrl}/auth/v1/user`, {
    headers: { apikey: anonKey, Authorization: `Bearer ${accessToken}` },
  });
  if (!userRes.ok) {
    return new Response(JSON.stringify({ error: "Sessão inválida" }), {
      status: 401, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
  const user = await userRes.json();

  const profileRes = await fetch(`${supabaseUrl}/rest/v1/profiles?id=eq.${user.id}&select=role,email`, {
    headers: { apikey: serviceKey, Authorization: `Bearer ${serviceKey}` },
  });
  const [profile] = await profileRes.json();
  if (!profile || profile.role !== "admin") {
    return new Response(JSON.stringify({ error: "Só Admin pode editar prompts de agente" }), {
      status: 403, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const svcHeaders = { apikey: serviceKey, Authorization: `Bearer ${serviceKey}`, "Content-Type": "application/json" };

  const agentRes = await fetch(`${supabaseUrl}/rest/v1/agents?name=eq.${agent_name}&select=id`, { headers: svcHeaders });
  const [agent] = await agentRes.json();
  if (!agent) {
    return new Response(JSON.stringify({ error: `Agente '${agent_name}' não encontrado` }), {
      status: 404, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const currentRes = await fetch(
    `${supabaseUrl}/rest/v1/agent_versions?agent_id=eq.${agent.id}&active=eq.true&select=id,version,model,execution_mode`,
    { headers: svcHeaders }
  );
  const [current] = await currentRes.json();
  const newVersion = current ? bumpVersion(current.version) : "v1.0";

  if (current) {
    await fetch(`${supabaseUrl}/rest/v1/agent_versions?id=eq.${current.id}`, {
      method: "PATCH", headers: svcHeaders, body: JSON.stringify({ active: false }),
    });
  }

  const insertVersionRes = await fetch(`${supabaseUrl}/rest/v1/agent_versions`, {
    method: "POST",
    headers: { ...svcHeaders, Prefer: "return=representation" },
    body: JSON.stringify({
      agent_id: agent.id,
      version: newVersion,
      model: model || current?.model || "claude-sonnet-4-6",
      execution_mode: current?.execution_mode || "subscription",
      active: true,
      created_by: user.id,
    }),
  });
  const [newVersionRow] = await insertVersionRes.json();

  await fetch(`${supabaseUrl}/rest/v1/agent_prompts`, {
    method: "POST",
    headers: svcHeaders,
    body: JSON.stringify({ agent_version_id: newVersionRow.id, system_prompt }),
  });

  await fetch(`${supabaseUrl}/rest/v1/activity_log`, {
    method: "POST",
    headers: svcHeaders,
    body: JSON.stringify({
      user_id: user.id, user_email: profile.email, action: "edit_prompt",
      details: { agent_name, new_version: newVersion },
    }),
  });

  return new Response(JSON.stringify({ ok: true, version: newVersion }), {
    status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
});
