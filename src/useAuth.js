import { useState, useEffect, useCallback } from "react";
import config from "./config";

const TOKEN_KEY = "ward_gtoken";
const USER_KEY  = "ward_guser";
const SCOPES    = "https://www.googleapis.com/auth/spreadsheets email profile openid";

export function useAuth() {
  const [user,    setUser]    = useState(() => {
    try {
      const u = JSON.parse(sessionStorage.getItem(USER_KEY));
      // Discard if already expired
      if (u && u.expiresAt && Date.now() > u.expiresAt) {
        sessionStorage.removeItem(USER_KEY); sessionStorage.removeItem(TOKEN_KEY);
        return null;
      }
      return u;
    } catch { return null; }
  });
  const [token,   setToken]   = useState(() => {
    try {
      const u = JSON.parse(sessionStorage.getItem(USER_KEY));
      if (u && u.expiresAt && Date.now() > u.expiresAt) return null;
      return sessionStorage.getItem(TOKEN_KEY) || null;
    } catch { return null; }
  });
  const [error,   setError]   = useState(null);
  const [loading, setLoading] = useState(false);

  const signIn = useCallback(() => {
    setError(null); setLoading(true);
    // Wait up to 5s for GSI script to finish loading before giving up
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
          sessionStorage.setItem(TOKEN_KEY, tr.access_token);
          sessionStorage.setItem(USER_KEY, JSON.stringify(userData));
          setToken(tr.access_token); setUser(userData); setLoading(false);
        } catch {
          setError("Could not verify your Google account. Please try again.");
          setLoading(false);
        }
      },
    });
    client.requestAccessToken({ prompt: "select_account" });
    }; // end doSignIn
    trySignIn();
  }, []);

  const signOut = useCallback(() => {
    if (token && window.google) window.google.accounts.oauth2.revoke(token, () => {});
    sessionStorage.removeItem(TOKEN_KEY); sessionStorage.removeItem(USER_KEY);
    setToken(null); setUser(null); setError(null);
  }, [token]);

  // Auto-refresh token 5 minutes before expiry
  useEffect(() => {
    if (!user) return;
    const msLeft = user.expiresAt - Date.now();
    if (msLeft <= 0) { signOut(); return; }
    const timer = setTimeout(() => {
      if (!window.google) return;
      const client = window.google.accounts.oauth2.initTokenClient({
        client_id: config.GOOGLE_CLIENT_ID, scope: SCOPES, prompt: "",
        callback: (tr) => {
          if (!tr.error) {
            const u = { ...user, expiresAt: Date.now() + (tr.expires_in - 60) * 1000 };
            sessionStorage.setItem(TOKEN_KEY, tr.access_token);
            sessionStorage.setItem(USER_KEY, JSON.stringify(u));
            setToken(tr.access_token); setUser(u);
          } else signOut();
        },
      });
      client.requestAccessToken({ prompt: "" });
    }, Math.max(msLeft - 5 * 60 * 1000, 1000));
    return () => clearTimeout(timer);
  }, [user, signOut]);

  return { user, token, error, loading, signIn, signOut };
}
