import { useState, useEffect, useCallback, useRef } from "react";
import config from "./config";

const TOKEN_KEY = "ward_gtoken";
const USER_KEY  = "ward_guser";
const SCOPES    = "https://www.googleapis.com/auth/spreadsheets email profile openid";

function silentRefresh(user, onSuccess, onFailure) {
  if (!window.google) { onFailure(); return; }
  const client = window.google.accounts.oauth2.initTokenClient({
    client_id: config.GOOGLE_CLIENT_ID,
    scope: SCOPES,
    prompt: "",
    callback: (tr) => {
      if (!tr.error) {
        const u = { ...user, expiresAt: Date.now() + (tr.expires_in - 60) * 1000 };
        localStorage.setItem(TOKEN_KEY, tr.access_token);
        localStorage.setItem(USER_KEY, JSON.stringify(u));
        onSuccess(tr.access_token, u);
      } else {
        onFailure();
      }
    },
  });
  client.requestAccessToken({ prompt: "" });
}

export function useAuth() {
  const [user,    setUser]    = useState(() => {
    try {
      const u = JSON.parse(localStorage.getItem(USER_KEY));
      if (u && u.expiresAt && Date.now() > u.expiresAt) {
        localStorage.removeItem(USER_KEY); localStorage.removeItem(TOKEN_KEY);
        return null;
      }
      return u;
    } catch { return null; }
  });
  const [token,   setToken]   = useState(() => {
    try {
      const u = JSON.parse(localStorage.getItem(USER_KEY));
      if (u && u.expiresAt && Date.now() > u.expiresAt) return null;
      return localStorage.getItem(TOKEN_KEY) || null;
    } catch { return null; }
  });
  const [error,   setError]   = useState(null);
  const [loading, setLoading] = useState(false);
  const refreshingRef = useRef(false);

  const signOut = useCallback(() => {
    if (token && window.google) window.google.accounts.oauth2.revoke(token, () => {});
    localStorage.removeItem(TOKEN_KEY); localStorage.removeItem(USER_KEY);
    setToken(null); setUser(null); setError(null);
  }, [token]);

  const signIn = useCallback(() => {
    setError(null); setLoading(true);
    let attempts = 0;
    const trySignIn = () => {
      if (!window.google) {
        attempts++;
        if (attempts > 50) {
          setError("Google Sign-In could not load. Check your internet connection and refresh.");
          setLoading(false); return;
        }
        return setTimeout(trySignIn, 100);
      }
      doSignIn();
    };
    const doSignIn = () => {
      const client = window.google.accounts.oauth2.initTokenClient({
        client_id: config.GOOGLE_CLIENT_ID,
        scope: SCOPES,
        callback: async (tr) => {
          if (tr.error) { setError("Sign-in failed: " + tr.error); setLoading(false); return; }
          try {
            const profile = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
              headers: { Authorization: `Bearer ${tr.access_token}` }
            }).then(r => r.json());
            const email   = profile.email?.toLowerCase();
            const allowed = [
              ...(config.ALLOWED_EMAILS||[]),
              ...(config.WARD_COUNCIL_EMAILS||[]),
            ].map(e => e.toLowerCase());
            if (!allowed.includes(email)) {
              setError(`Access denied. ${profile.email} is not authorised. Contact your ward administrator.`);
              setLoading(false); return;
            }
            const userData = { email: profile.email, name: profile.name, picture: profile.picture,
              expiresAt: Date.now() + (tr.expires_in - 60) * 1000 };
            localStorage.setItem(TOKEN_KEY, tr.access_token);
            localStorage.setItem(USER_KEY, JSON.stringify(userData));
            setToken(tr.access_token); setUser(userData); setLoading(false);
          } catch {
            setError("Could not verify your Google account. Please try again.");
            setLoading(false);
          }
        },
      });
      client.requestAccessToken({ prompt: "select_account" });
    };
    trySignIn();
  }, []);

  // ── Token refresh strategy ──────────────────────────────────────────────────
  // Two triggers:
  // 1. setTimeout fires 5 min before expiry (works when app is in foreground)
  // 2. visibilitychange fires when user returns to the app (catches iOS background kills)
  // Both check expiry before acting so they don't double-refresh.

  const maybeRefresh = useCallback(() => {
    if (!user || refreshingRef.current) return;
    const msLeft = user.expiresAt - Date.now();
    // Refresh if less than 10 minutes remain (generous window for iOS wake-up lag)
    if (msLeft > 10 * 60 * 1000) return;
    if (msLeft <= 0) { signOut(); return; }
    refreshingRef.current = true;
    silentRefresh(
      user,
      (newToken, newUser) => { setToken(newToken); setUser(newUser); refreshingRef.current = false; },
      () => { refreshingRef.current = false; signOut(); }
    );
  }, [user, signOut]);

  // Timer-based refresh (foreground)
  useEffect(() => {
    if (!user) return;
    const msLeft = user.expiresAt - Date.now();
    if (msLeft <= 0) { signOut(); return; }
    // Fire 5 min before expiry
    const delay = Math.max(msLeft - 5 * 60 * 1000, 1000);
    const timer = setTimeout(maybeRefresh, delay);
    return () => clearTimeout(timer);
  }, [user, signOut, maybeRefresh]);

  // Visibility-based refresh (catches iOS background + tab switching)
  useEffect(() => {
    if (!user) return;
    const handler = () => {
      if (document.visibilityState === "visible") maybeRefresh();
    };
    document.addEventListener("visibilitychange", handler);
    return () => document.removeEventListener("visibilitychange", handler);
  }, [user, maybeRefresh]);

  return { user, token, error, loading, signIn, signOut };
}
