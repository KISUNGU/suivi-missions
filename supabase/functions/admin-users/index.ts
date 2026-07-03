// ════════════════════════════════════════════════════════════════════════════
// PNDA-SE — Edge Function admin-users
// ════════════════════════════════════════════════════════════════════════════
// Gère la création, la modification, la réinitialisation de mot de passe et
// la suppression des comptes utilisateurs (Supabase Auth + table `profiles`).
//
// Ces opérations nécessitent la clé service_role (createUser, updateUserById,
// deleteUser via l'Admin API), qui ne doit jamais être exposée au navigateur.
// Cette fonction Edge est le seul endroit où cette clé est utilisée : elle
// tourne côté serveur, vérifie que l'appelant est authentifié ET a le rôle
// super_admin (RNSE) dans `profiles`, puis exécute l'action demandée.
//
// Appelée depuis superadmin.html via fetch(), avec le token d'accès de
// l'utilisateur connecté en en-tête Authorization.
//
// Déploiement (Supabase CLI) :
//   supabase functions deploy admin-users --project-ref <PROJECT_REF>
// ════════════════════════════════════════════════════════════════════════════

import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "jsr:@supabase/supabase-js@2";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_ROLE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;

const CORS_HEADERS: Record<string, string> = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};

function jsonResponse(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...CORS_HEADERS, "Content-Type": "application/json" },
  });
}

const VALID_ROLES = ["province", "comptable", "admin", "super_admin"];

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response("ok", { headers: CORS_HEADERS });
  }

  if (req.method !== "POST") {
    return jsonResponse({ error: "Méthode non supportée." }, 405);
  }

  const authHeader = req.headers.get("Authorization") || "";
  const jwt = authHeader.replace(/^Bearer\s+/i, "");
  if (!jwt) {
    return jsonResponse({ error: "Authentification requise." }, 401);
  }

  const admin = createClient(SUPABASE_URL, SERVICE_ROLE_KEY);

  // Vérifier le JWT de l'appelant et son rôle super_admin
  const { data: callerData, error: callerErr } = await admin.auth.getUser(jwt);
  if (callerErr || !callerData?.user) {
    return jsonResponse({ error: "Session invalide ou expirée." }, 401);
  }

  const { data: callerProfile, error: profileErr } = await admin
    .from("profiles")
    .select("role")
    .eq("id", callerData.user.id)
    .single();

  if (profileErr || callerProfile?.role !== "super_admin") {
    return jsonResponse({ error: "Accès réservé au super-administrateur (RNSE)." }, 403);
  }

  let body: Record<string, unknown>;
  try {
    body = await req.json();
  } catch {
    return jsonResponse({ error: "Corps de requête JSON invalide." }, 400);
  }

  const action = String(body.action || "");

  try {
    if (action === "create") {
      const email = String(body.email || "").trim().toLowerCase();
      const password = String(body.password || "");
      const role = String(body.role || "");
      const provinces = (body.provinces ?? null) as string[] | null;
      const display_name = (body.display_name ?? null) as string | null;

      if (!email || !password || !VALID_ROLES.includes(role)) {
        return jsonResponse({ error: "Email, mot de passe et rôle valide sont requis." }, 400);
      }
      if (password.length < 6) {
        return jsonResponse({ error: "Le mot de passe doit contenir au moins 6 caractères." }, 400);
      }

      const { data: created, error: createErr } = await admin.auth.admin.createUser({
        email,
        password,
        email_confirm: true,
      });
      if (createErr || !created?.user) {
        return jsonResponse({ error: createErr?.message || "Création du compte impossible." }, 400);
      }

      const { error: insErr } = await admin.from("profiles").insert({
        id: created.user.id,
        email,
        role,
        provinces,
        display_name,
      });
      if (insErr) {
        await admin.auth.admin.deleteUser(created.user.id);
        return jsonResponse({ error: insErr.message }, 400);
      }

      return jsonResponse({ success: true, id: created.user.id });
    }

    if (action === "update") {
      const id = String(body.id || "");
      if (!id) return jsonResponse({ error: "Identifiant utilisateur requis." }, 400);

      const patch: Record<string, unknown> = { updated_at: new Date().toISOString() };
      if (body.role !== undefined) {
        if (!VALID_ROLES.includes(String(body.role))) {
          return jsonResponse({ error: "Rôle invalide." }, 400);
        }
        patch.role = body.role;
      }
      if (body.provinces !== undefined) patch.provinces = body.provinces;
      if (body.display_name !== undefined) patch.display_name = body.display_name;
      if (body.is_active !== undefined) patch.is_active = body.is_active;

      const { error: updErr } = await admin.from("profiles").update(patch).eq("id", id);
      if (updErr) return jsonResponse({ error: updErr.message }, 400);

      if (body.is_active === false) {
        await admin.auth.admin.updateUserById(id, { ban_duration: "876000h" });
      } else if (body.is_active === true) {
        await admin.auth.admin.updateUserById(id, { ban_duration: "none" });
      }

      return jsonResponse({ success: true });
    }

    if (action === "reset_password") {
      const id = String(body.id || "");
      const password = String(body.password || "");
      if (!id || password.length < 6) {
        return jsonResponse({ error: "Identifiant et mot de passe (6 caractères min.) requis." }, 400);
      }
      const { error: pwErr } = await admin.auth.admin.updateUserById(id, { password });
      if (pwErr) return jsonResponse({ error: pwErr.message }, 400);
      return jsonResponse({ success: true });
    }

    if (action === "delete") {
      const id = String(body.id || "");
      if (!id) return jsonResponse({ error: "Identifiant utilisateur requis." }, 400);
      if (id === callerData.user.id) {
        return jsonResponse({ error: "Impossible de supprimer votre propre compte." }, 400);
      }
      await admin.from("profiles").delete().eq("id", id);
      const { error: delErr } = await admin.auth.admin.deleteUser(id);
      if (delErr) return jsonResponse({ error: delErr.message }, 400);
      return jsonResponse({ success: true });
    }

    return jsonResponse({ error: "Action inconnue: " + action }, 400);
  } catch (e) {
    return jsonResponse({ error: e instanceof Error ? e.message : String(e) }, 500);
  }
});
