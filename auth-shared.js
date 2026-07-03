// ════════════════════════════════════════════════════════════════════════════
// PNDA-SE — Authentification partagée (Supabase Auth)
// ════════════════════════════════════════════════════════════════════════════
// Ce fichier fournit un écran de connexion commun aux trois pages de
// l'application (index.html, admin.html, superadmin.html) et gère la
// création / le rafraîchissement de la session locale (PNDA_SESSION) à
// partir de la session Supabase Auth réelle + de la table `profiles`.
//
// Utilisation dans chaque page :
//   const session = await PndaAuth.hydrateSession(db, ['admin','comptable']);
//   if (!session) return; // l'écran de connexion est affiché, on attend l'utilisateur
//
// PndaAuth.logout(db) déconnecte proprement (Supabase Auth + session locale).
// ────────────────────────────────────────────────────────────────────────────

const PndaAuth = (() => {
  const ROLE_LABELS = {
    province: 'Missionnaire (province)',
    comptable: 'Comptable',
    admin: 'Administrateur / RAF',
    super_admin: 'Super-administrateur (RNSE)'
  };

  function injectStyles() {
    if (document.getElementById('pnda-auth-styles')) return;
    const style = document.createElement('style');
    style.id = 'pnda-auth-styles';
    style.textContent = `
      #pnda-auth-overlay {
        position: fixed; inset: 0; z-index: 2000;
        display: flex; align-items: center; justify-content: center;
        background: radial-gradient(circle at top, #1d4c34 0%, #0d2318 65%, #081712 100%);
        padding: 20px;
      }
      #pnda-auth-overlay .auth-card {
        width: 100%; max-width: 400px;
        background: #fffdf7;
        border-radius: 14px;
        box-shadow: 0 24px 60px rgba(0,0,0,0.45);
        overflow: hidden;
        border-top: 4px solid #d3a52c;
      }
      #pnda-auth-overlay .auth-card-header {
        padding: 28px 32px 18px;
        text-align: center;
      }
      #pnda-auth-overlay .auth-card-header img { height: 56px; margin-bottom: 10px; }
      #pnda-auth-overlay .auth-card-header h1 {
        font-family: 'Playfair Display', serif;
        font-size: 1.25rem; color: #123524; margin: 0 0 2px;
      }
      #pnda-auth-overlay .auth-card-header p {
        font-size: 0.75rem; letter-spacing: 0.5px; color: #6b7d72;
        text-transform: uppercase; margin: 0;
      }
      #pnda-auth-overlay .auth-card-body { padding: 8px 32px 32px; }
      #pnda-auth-overlay .form-label {
        font-size: 0.72rem; font-weight: 700; letter-spacing: 0.6px;
        text-transform: uppercase; color: #1d4c34;
      }
      #pnda-auth-overlay .btn-pnda-auth {
        background: #1d4c34; border-color: #1d4c34; color: #fff;
        font-weight: 600; letter-spacing: 0.3px;
      }
      #pnda-auth-overlay .btn-pnda-auth:hover { background: #123524; border-color: #123524; color: #fff; }
      #pnda-auth-overlay .auth-alert { font-size: 0.85rem; }
      #pnda-auth-overlay .auth-footer {
        text-align: center; font-size: 0.7rem; color: rgba(255,255,255,0.55);
        margin-top: 18px; letter-spacing: 0.4px;
      }
    `;
    document.head.appendChild(style);
  }

  function renderLoginMarkup(deniedMessage) {
    return `
      <div class="auth-card">
        <div class="auth-card-header">
          <img src="logo_pnda.png" alt="PNDA-SE">
          <h1>PNDA-SE</h1>
          <p>Suivi &amp; évaluation des missions</p>
        </div>
        <div class="auth-card-body">
          <div id="pnda-auth-error" class="alert alert-danger auth-alert py-2 px-3 mb-3 ${deniedMessage ? '' : 'd-none'}">${deniedMessage || ''}</div>
          <form id="pnda-auth-form" novalidate>
            <div class="mb-3">
              <label class="form-label" for="pnda-auth-email">Adresse email</label>
              <input type="email" class="form-control" id="pnda-auth-email" autocomplete="username" required placeholder="nom@pnda.cd">
            </div>
            <div class="mb-3">
              <label class="form-label" for="pnda-auth-password">Mot de passe</label>
              <input type="password" class="form-control" id="pnda-auth-password" autocomplete="current-password" required placeholder="••••••••">
            </div>
            <button type="submit" class="btn btn-pnda-auth w-100 py-2" id="pnda-auth-submit">
              <i class="bi bi-box-arrow-in-right me-1"></i>Se connecter
            </button>
          </form>
        </div>
      </div>
      <div class="auth-footer">Programme National de Développement Agricole · RD Congo</div>
    `;
  }

  function showLoginScreen(deniedMessage) {
    injectStyles();
    let overlay = document.getElementById('pnda-auth-overlay');
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = 'pnda-auth-overlay';
      document.body.appendChild(overlay);
    }
    overlay.innerHTML = renderLoginMarkup(deniedMessage);
    overlay.style.display = 'flex';
    document.getElementById('pnda-auth-email')?.focus();
    return overlay;
  }

  function hideLoginScreen() {
    const overlay = document.getElementById('pnda-auth-overlay');
    if (overlay) overlay.style.display = 'none';
  }

  function setError(message) {
    const el = document.getElementById('pnda-auth-error');
    if (!el) return;
    if (message) {
      el.textContent = message;
      el.classList.remove('d-none');
    } else {
      el.classList.add('d-none');
    }
  }

  async function fetchProfile(db, userId) {
    const { data, error } = await db.from('profiles').select('*').eq('id', userId).single();
    if (error) throw error;
    return data;
  }

  function buildLocalSession(profile) {
    const provinces = Array.isArray(profile.provinces) ? profile.provinces : null;
    return {
      user: profile.email,
      role: profile.role,
      displayName: profile.display_name || profile.email,
      provinces,
      province: (provinces && provinces.length === 1) ? provinces[0] : null
    };
  }

  function persistSession(session) {
    localStorage.setItem('PNDA_SESSION', JSON.stringify(session));
  }

  function getSession() {
    try { return JSON.parse(localStorage.getItem('PNDA_SESSION') || 'null'); }
    catch { return null; }
  }

  async function accessTokenFor(db) {
    const { data } = await db.auth.getSession();
    return data?.session?.access_token || null;
  }

  /**
   * Vérifie/établit une session valide pour l'un des rôles autorisés.
   * Affiche un écran de connexion tant que ce n'est pas le cas.
   * Résout avec l'objet session locale ({user, role, province, provinces, displayName})
   * une fois la connexion réussie et le rôle validé.
   */
  function hydrateSession(db, allowedRoles) {
    return new Promise((resolve) => {
      let settled = false;

      async function attemptFromExistingAuthSession() {
        const { data: { session: authSession } } = await db.auth.getSession();
        if (!authSession) {
          showLoginScreen();
          return;
        }
        try {
          const profile = await fetchProfile(db, authSession.user.id);
          if (!profile.is_active || !allowedRoles.includes(profile.role)) {
            await db.auth.signOut();
            localStorage.removeItem('PNDA_SESSION');
            showLoginScreen(`Ce compte (${ROLE_LABELS[profile.role] || profile.role}) n'a pas accès à cette page.`);
            return;
          }
          const localSession = buildLocalSession(profile);
          persistSession(localSession);
          hideLoginScreen();
          settled = true;
          resolve(localSession);
        } catch (e) {
          console.error('Erreur de chargement du profil:', e);
          await db.auth.signOut();
          localStorage.removeItem('PNDA_SESSION');
          showLoginScreen("Impossible de charger votre profil. Contactez le RNSE.");
        }
      }

      function bindFormHandler() {
        const form = document.getElementById('pnda-auth-form');
        if (!form || form.dataset.bound === '1') return;
        form.dataset.bound = '1';
        form.addEventListener('submit', async (event) => {
          event.preventDefault();
          setError('');
          const email = document.getElementById('pnda-auth-email').value.trim().toLowerCase();
          const password = document.getElementById('pnda-auth-password').value;
          const submitBtn = document.getElementById('pnda-auth-submit');
          submitBtn.disabled = true;
          submitBtn.innerHTML = '<span class="spinner-border spinner-border-sm me-1"></span>Connexion…';
          try {
            const { data, error } = await db.auth.signInWithPassword({ email, password });
            if (error) throw error;
            const profile = await fetchProfile(db, data.user.id);
            if (!profile.is_active || !allowedRoles.includes(profile.role)) {
              await db.auth.signOut();
              setError(`Ce compte (${ROLE_LABELS[profile.role] || profile.role}) n'a pas accès à cette page.`);
              return;
            }
            const localSession = buildLocalSession(profile);
            persistSession(localSession);
            hideLoginScreen();
            if (!settled) { settled = true; resolve(localSession); }
          } catch (e) {
            console.error('Erreur de connexion:', e);
            const msg = /invalid login credentials/i.test(e?.message || '')
              ? 'Email ou mot de passe incorrect.'
              : (e?.message || 'Erreur de connexion.');
            setError(msg);
          } finally {
            submitBtn.disabled = false;
            submitBtn.innerHTML = '<i class="bi bi-box-arrow-in-right me-1"></i>Se connecter';
          }
        });
      }

      // Ré-attache le handler du formulaire de connexion dès qu'il est injecté au DOM
      const observer = new MutationObserver(() => bindFormHandler());
      observer.observe(document.body, { childList: true, subtree: true });

      attemptFromExistingAuthSession();
    });
  }

  async function logout(db) {
    try { await db.auth.signOut(); } catch (e) { console.warn('signOut error', e); }
    localStorage.removeItem('PNDA_SESSION');
    window.location.reload();
  }

  return { hydrateSession, getSession, logout, accessTokenFor, ROLE_LABELS };
})();
