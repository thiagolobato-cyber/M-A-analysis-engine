// Convida alguém por e-mail (usa o serviço de e-mail do próprio Supabase —
// a pessoa recebe um link, clica, define a senha, e já entra com o papel
// escolhido aqui). Só Admin pode chamar isso.
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, content-type, apikey",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: corsHeaders });
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), {
      status: 405, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const { email, role } = await req.json().catch(() => ({}));
  if (!email || !["admin", "analyst"].includes(role)) {
    return new Response(JSON.stringify({ error: "email e role ('admin'|'analyst') são obrigatórios" }), {
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
  const caller = await userRes.json();

  const profileRes = await fetch(`${supabaseUrl}/rest/v1/profiles?id=eq.${caller.id}&select=role,email`, {
    headers: { apikey: serviceKey, Authorization: `Bearer ${serviceKey}` },
  });
  const [callerProfile] = await profileRes.json();
  if (!callerProfile || callerProfile.role !== "admin") {
    return new Response(JSON.stringify({ error: "Só Admin pode convidar pessoas" }), {
      status: 403, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  const supabaseAdmin = createClient(supabaseUrl, serviceKey);
  const { data, error } = await supabaseAdmin.auth.admin.inviteUserByEmail(email);
  if (error) {
    return new Response(JSON.stringify({ error: error.message }), {
      status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }

  // O trigger handle_new_user já cria o profile com role='analyst' por
  // padrão — se o admin escolheu 'admin', atualiza aqui.
  if (role === "admin" && data.user) {
    await fetch(`${supabaseUrl}/rest/v1/profiles?id=eq.${data.user.id}`, {
      method: "PATCH",
      headers: { apikey: serviceKey, Authorization: `Bearer ${serviceKey}`, "Content-Type": "application/json" },
      body: JSON.stringify({ role: "admin" }),
    });
  }

  await fetch(`${supabaseUrl}/rest/v1/activity_log`, {
    method: "POST",
    headers: { apikey: serviceKey, Authorization: `Bearer ${serviceKey}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      user_id: caller.id, user_email: callerProfile.email, action: "invite_user",
      details: { invited_email: email, role },
    }),
  });

  return new Response(JSON.stringify({ ok: true }), {
    status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
});
