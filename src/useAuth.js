import { useState, useCallback } from "react";
import config from "./config";

const USER_KEY = "ward_user";

function getAllowedEmails() {
  return [
    ...(config.ALLOWED_EMAILS      || []),
    ...(config.WARD_COUNCIL_EMAILS || []),
  ].map(e => e.toLowerCase());
}

function loadSession() {
  try {
    const u = JSON.parse(localStorage.getItem(USER_KEY));
    return u || null;
  } catch { return null; }
}

export function useAuth() {
  const [user, setUser] = useState(() => loadSession());
  const [error, setError] = useState(null);

  const submitEmail = useCallback((email) => {
    const trimmed = email.trim().toLowerCase();
    if (!trimmed) { setError("Please enter your email address."); return; }
    if (!getAllowedEmails().includes(trimmed)) {
      setError(`${trimmed} is not authorised. Contact your ward administrator.`);
      return;
    }
    setError(null);
    const userData = { email: trimmed, name: trimmed.split("@")[0] };
    localStorage.setItem(USER_KEY, JSON.stringify(userData));
    setUser(userData);
  }, []);

  const signOut = useCallback(() => {
    localStorage.removeItem(USER_KEY);
    setUser(null);
    setError(null);
  }, []);

  return { user, error, submitEmail, signOut };
}
