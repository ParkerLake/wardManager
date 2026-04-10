import { useState, useEffect, useCallback, useRef } from "react";
import config from "./config";

const TOKEN_KEY  = "ward_gtoken";
const USER_KEY   = "ward_guser";
const SCOPES     = "https://www.googleapis.com/auth/spreadsheets";

// ── Helpers ──────────────────────────────────────────────────────────────────

function getAllowedEmails() {
  return [
    ...(config.ALLOWED_EMAILS      || []),
    ...(config.WARD_COUNCIL_EMAILS || []),
  ].map(e => e.toLowerCase());
}

function loadSession() {
  try {
    const u = JSON.parse(localStorage.getItem(USER_KEY));
    if (!u) return null;
    const token = localStorage.getItem(TOKEN_KEY);
    const tokenExpiry = parseInt(localStorage.getItem(TOKEN_KEY + "_expiry") || "0");
    return {
      user: u,
      token: Date.now() < tokenExpiry ? token : null,
    };
  } catch { return null; }
}

function saveEmailSession(email) {
  const userData = { email, name: email.split("@")[0], picture: null };
  localStorage.setItem(USER_KEY, JSON.stringify(userData));
  return userData;
}

function saveToken(token, expiresIn) {
  localStorage.setItem(TOKEN_KEY, token);
  localStorage.setItem(TOKEN_KEY + "_expiry", String(Date.now() + (expiresIn - 60) * 1000));
}

function clearToken() {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(TOKEN_KEY + "_expiry");
}

// ── Hook ─────────────────────────────────────────────────────────────────────

export function useAuth() {
  const session = loadSession();
  const [user,    setUser]    = useState(session?.user  || null);
  const [token,   setToken]   = useState(session?.token || null);
  const [error,   setError]   = useState(null);
  const [loading, setLoading] = useState(false);
  const [stage,   setStage]   = useState(session?.user ? "done" : "email_entry");
  const refreshingRef = useRef(false);

  // ── acquireToken — always runs in background, never blocks UI ────────────
  const acquireToken = useCallback((hint) => {
    if (refreshingRef.current) return;
    if (!window.google) {
      const wait = setInterval(() => {
        if (window.google) { clearInterval(wait); acquireToken(hint); }
      }, 200);
      return;
    }
    refreshingRef.current = true;
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: config.GOOGLE_CLIENT_ID,
      scope: SCOPES,
      prompt: "",
      login_hint: hint || "",
      callback: (tr) => {
        refreshingRef.current = false;
        if (!tr.error) {
          saveToken(tr.access_token, tr.expires_in);
          setToken(tr.access_token);
          setStage("done");
        } else {
          // Silent failed — show connect banner inside the app (non-blocking)
          setStage("needs_google");
        }
      },
    });
    client.requestAccessToken({ prompt: "" });
  }, []); // eslint-disable-line

  // ── Step 1: Email gate — no Google involved ──────────────────────────────
  const submitEmail = useCallback((email) => {
    const trimmed = email.trim().toLowerCase();
    if (!trimmed) { setError("Please enter your email address."); return; }
    if (!getAllowedEmails().includes(trimmed)) {
      setError(`${trimmed} is not authorised. Contact your ward administrator.`);
      return;
    }
    setError(null);
    const userData = saveEmailSession(trimmed);
    setUser(userData);
    setStage("done");
    // Silently acquire Google token in background for Sheets access
    acquireToken(trimmed);
  }, [acquireToken]);

  // ── Explicit Google sign-in — called from banner inside the app ──────────
  const googleSignIn = useCallback(() => {
    if (!window.google) return;
    setLoading(true);
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: config.GOOGLE_CLIENT_ID,
      scope: SCOPES,
      login_hint: user?.email || "",
      callback: (tr) => {
        setLoading(false);
        if (!tr.error) {
          saveToken(tr.access_token, tr.expires_in);
          setToken(tr.access_token);
          setStage("done");
        }
      },
    });
    client.requestAccessToken({ prompt: "select_account" });
  }, [user]);

  // ── Sign out ──────────────────────────────────────────────────────────────
  const signOut = useCallback(() => {
    if (token && window.google) window.google.accounts.oauth2.revoke(token, () => {});
    localStorage.removeItem(USER_KEY);
    clearToken();
    setToken(null); setUser(null); setError(null);
    setStage("email_entry");
  }, [token]);

  // ── Token refresh — timer + visibilitychange (iOS) ───────────────────────
  const maybeRefresh = useCallback(() => {
    if (!user || refreshingRef.current) return;
    const expiry = parseInt(localStorage.getItem(TOKEN_KEY + "_expiry") || "0");
    const msLeft = expiry - Date.now();
    if (msLeft > 10 * 60 * 1000) return;
    acquireToken(user.email);
  }, [user, acquireToken]);

  useEffect(() => {
    if (!user || !token) return;
    const expiry = parseInt(localStorage.getItem(TOKEN_KEY + "_expiry") || "0");
    const msLeft = expiry - Date.now();
    if (msLeft <= 0) { clearToken(); setToken(null); return; }
    const timer = setTimeout(maybeRefresh, Math.max(msLeft - 5 * 60 * 1000, 1000));
    return () => clearTimeout(timer);
  }, [user, token, maybeRefresh]);

  useEffect(() => {
    if (!user) return;
    const handler = () => { if (document.visibilityState === "visible") maybeRefresh(); };
    document.addEventListener("visibilitychange", handler);
    return () => document.removeEventListener("visibilitychange", handler);
  }, [user, maybeRefresh]);

  return { user, token, error, loading, stage, submitEmail, googleSignIn, signOut };
}
