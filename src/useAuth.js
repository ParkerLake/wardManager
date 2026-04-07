import { useState, useEffect, useCallback, useRef } from "react";
import config from "./config";

const TOKEN_KEY   = "ward_gtoken";
const USER_KEY    = "ward_guser";
const SCOPES      = "https://www.googleapis.com/auth/spreadsheets email profile openid";

// ── Helpers ──────────────────────────────────────────────────────────────────

function getAllowedEmails() {
  return [
    ...(config.ALLOWED_EMAILS      || []),
    ...(config.WARD_COUNCIL_EMAILS || []),
  ].map(e => e.toLowerCase());
}

function isEmailAllowed(email) {
  return getAllowedEmails().includes(email.toLowerCase().trim());
}

function saveSession(token, userData) {
  localStorage.setItem(TOKEN_KEY, token);
  localStorage.setItem(USER_KEY, JSON.stringify(userData));
}

function loadSession() {
  try {
    const u = JSON.parse(localStorage.getItem(USER_KEY));
    if (!u || !u.expiresAt) return null;
    if (Date.now() > u.expiresAt) {
      localStorage.removeItem(USER_KEY);
      localStorage.removeItem(TOKEN_KEY);
      return null;
    }
    return { user: u, token: localStorage.getItem(TOKEN_KEY) };
  } catch { return null; }
}

// Attempt silent Google OAuth — no prompt, no popup, no warning
// Resolves with access_token if Google already has a session, rejects if not
function silentGoogleToken(hint) {
  return new Promise((resolve, reject) => {
    if (!window.google) { reject(new Error("GSI not loaded")); return; }
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: config.GOOGLE_CLIENT_ID,
      scope: SCOPES,
      prompt: "",
      login_hint: hint || "",
      callback: (tr) => {
        if (tr.error) reject(new Error(tr.error));
        else resolve(tr);
      },
    });
    client.requestAccessToken({ prompt: "" });
  });
}

// Explicit Google sign-in — shows the account picker / warning if needed
function explicitGoogleToken(hint) {
  return new Promise((resolve, reject) => {
    if (!window.google) { reject(new Error("GSI not loaded")); return; }
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: config.GOOGLE_CLIENT_ID,
      scope: SCOPES,
      login_hint: hint || "",
      callback: (tr) => {
        if (tr.error) reject(new Error(tr.error));
        else resolve(tr);
      },
    });
    client.requestAccessToken({ prompt: "select_account" });
  });
}

function buildUserData(email, name, picture, expiresIn) {
  return {
    email,
    name: name || email.split("@")[0],
    picture: picture || null,
    expiresAt: Date.now() + (expiresIn - 60) * 1000,
  };
}

// ── Hook ─────────────────────────────────────────────────────────────────────

export function useAuth() {
  const session = loadSession();
  const [user,    setUser]    = useState(session?.user  || null);
  const [token,   setToken]   = useState(session?.token || null);
  const [error,   setError]   = useState(null);
  const [loading, setLoading] = useState(false);
  // "email_entry"   — user is typing their email
  // "acquiring"     — silently getting Google token
  // "needs_google"  — silent failed, show explicit sign-in button
  // "done"          — authenticated
  const [stage, setStage]     = useState(session ? "done" : "email_entry");
  const [pendingEmail, setPendingEmail] = useState("");
  const refreshingRef = useRef(false);

  const signOut = useCallback(() => {
    if (token && window.google) window.google.accounts.oauth2.revoke(token, () => {});
    localStorage.removeItem(TOKEN_KEY);
    localStorage.removeItem(USER_KEY);
    setToken(null); setUser(null); setError(null);
    setStage("email_entry"); setPendingEmail("");
  }, [token]);

  // Step 1: user submits email
  const submitEmail = useCallback(async (email) => {
    const trimmed = email.trim().toLowerCase();
    if (!trimmed) { setError("Please enter your email address."); return; }
    if (!isEmailAllowed(trimmed)) {
      setError(`${trimmed} is not authorised. Contact your ward administrator.`);
      return;
    }
    setError(null);
    setPendingEmail(trimmed);
    setLoading(true);
    setStage("acquiring");

    // Wait for GSI to load (up to 5s)
    let waited = 0;
    while (!window.google && waited < 50) {
      await new Promise(r => setTimeout(r, 100));
      waited++;
    }
    if (!window.google) {
      setLoading(false);
      setStage("needs_google");
      return;
    }

    // Try silent token acquisition first
    try {
      const tr = await silentGoogleToken(trimmed);
      await finishAuth(tr, trimmed);
    } catch {
      // Silent failed — user needs to click the Google button once
      setLoading(false);
      setStage("needs_google");
    }
  }, []); // eslint-disable-line

  // Step 2 (if needed): explicit Google sign-in
  const googleSignIn = useCallback(async () => {
    setError(null);
    setLoading(true);
    try {
      const tr = await explicitGoogleToken(pendingEmail);
      await finishAuth(tr, pendingEmail);
    } catch (e) {
      setError("Sign-in failed. Please try again.");
      setLoading(false);
      setStage("needs_google");
    }
  }, [pendingEmail]); // eslint-disable-line

  async function finishAuth(tr, email) {
    try {
      // Fetch profile to confirm identity
      const profile = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
        headers: { Authorization: `Bearer ${tr.access_token}` }
      }).then(r => r.json());

      const confirmedEmail = profile.email?.toLowerCase();

      // Double-check the Google account matches the entered email
      if (confirmedEmail !== email.toLowerCase()) {
        setError(`Signed in as ${profile.email} but you entered ${email}. Please use the correct Google account.`);
        setLoading(false);
        setStage("needs_google");
        return;
      }

      const userData = buildUserData(profile.email, profile.name, profile.picture, tr.expires_in);
      saveSession(tr.access_token, userData);
      setToken(tr.access_token);
      setUser(userData);
      setLoading(false);
      setStage("done");
    } catch {
      setError("Could not verify your account. Please try again.");
      setLoading(false);
      setStage("needs_google");
    }
  }

  // ── Token refresh (same as before — timer + visibility) ──────────────────
  const maybeRefresh = useCallback(() => {
    if (!user || refreshingRef.current) return;
    const msLeft = user.expiresAt - Date.now();
    if (msLeft > 10 * 60 * 1000) return;
    if (msLeft <= 0) { signOut(); return; }
    refreshingRef.current = true;
    if (!window.google) { signOut(); return; }
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: config.GOOGLE_CLIENT_ID,
      scope: SCOPES,
      prompt: "",
      login_hint: user.email,
      callback: (tr) => {
        refreshingRef.current = false;
        if (!tr.error) {
          const u = buildUserData(user.email, user.name, user.picture, tr.expires_in);
          saveSession(tr.access_token, u);
          setToken(tr.access_token); setUser(u);
        } else { signOut(); }
      },
    });
    client.requestAccessToken({ prompt: "" });
  }, [user, signOut]);

  useEffect(() => {
    if (!user) return;
    const msLeft = user.expiresAt - Date.now();
    if (msLeft <= 0) { signOut(); return; }
    const timer = setTimeout(maybeRefresh, Math.max(msLeft - 5 * 60 * 1000, 1000));
    return () => clearTimeout(timer);
  }, [user, signOut, maybeRefresh]);

  useEffect(() => {
    if (!user) return;
    const handler = () => { if (document.visibilityState === "visible") maybeRefresh(); };
    document.addEventListener("visibilitychange", handler);
    return () => document.removeEventListener("visibilitychange", handler);
  }, [user, maybeRefresh]);

  return { user, token, error, loading, stage, pendingEmail, submitEmail, googleSignIn, signOut };
}
