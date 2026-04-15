import { useState, useCallback } from "react";
import config from "./config";

const USER_KEY = "ward_user";

function isAdmin(email) {
  return (config.ALLOWED_EMAILS || []).map(e => e.toLowerCase()).includes(email.toLowerCase());
}

function isWardCouncil(email) {
  return (config.WARD_COUNCIL_EMAILS || []).map(e => e.toLowerCase()).includes(email.toLowerCase());
}

function getAllowedEmails() {
  return [
    ...(config.ALLOWED_EMAILS      || []),
    ...(config.WARD_COUNCIL_EMAILS || []),
  ].map(e => e.toLowerCase());
}

async function sha256(text) {
  const buf = await crypto.subtle.digest(
    "SHA-256",
    new TextEncoder().encode(text)
  );
  return Array.from(new Uint8Array(buf))
    .map(b => b.toString(16).padStart(2, "0"))
    .join("");
}

function loadSession() {
  try {
    const u = JSON.parse(localStorage.getItem(USER_KEY));
    return u || null;
  } catch { return null; }
}

export function useAuth() {
  const [user,         setUser]         = useState(() => loadSession());
  const [error,        setError]        = useState(null);
  const [stage,        setStage]        = useState(() => loadSession() ? "done" : "email_entry");
  const [pendingEmail, setPendingEmail] = useState("");

  // Step 1: email check
  const submitEmail = useCallback((email) => {
    const trimmed = email.trim().toLowerCase();
    if (!trimmed) { setError("Please enter your email address."); return; }
    if (!getAllowedEmails().includes(trimmed)) {
      setError(`${trimmed} is not authorised. Contact your ward administrator.`);
      return;
    }
    setError(null);
    setPendingEmail(trimmed);
    // Check if a password hash is configured for this user's role
    const needsPassword = isAdmin(trimmed)
      ? !!config.ADMIN_PASSWORD_HASH
      : !!config.WC_PASSWORD_HASH;
    if (!needsPassword) {
      const userData = { email: trimmed, name: trimmed.split("@")[0] };
      localStorage.setItem(USER_KEY, JSON.stringify(userData));
      setUser(userData);
      setStage("done");
    } else {
      setStage("password_entry");
    }
  }, []);

  // Step 2: password check — verifies against the correct hash for their role
  const submitPassword = useCallback(async (password) => {
    if (!password) { setError("Please enter the password."); return; }
    setError(null);
    const hash = await sha256(password.trim());
    const expectedHash = isAdmin(pendingEmail)
      ? config.ADMIN_PASSWORD_HASH
      : config.WC_PASSWORD_HASH;
    if (hash !== expectedHash) {
      setError("Incorrect password. Please try again.");
      return;
    }
    const userData = { email: pendingEmail, name: pendingEmail.split("@")[0] };
    localStorage.setItem(USER_KEY, JSON.stringify(userData));
    setUser(userData);
    setStage("done");
  }, [pendingEmail]);

  const signOut = useCallback(() => {
    localStorage.removeItem(USER_KEY);
    setUser(null); setError(null);
    setPendingEmail("");
    setStage("email_entry");
  }, []);

  const backToEmail = useCallback(() => {
    setStage("email_entry");
    setPendingEmail("");
    setError(null);
  }, []);

  return { user, error, stage, pendingEmail, submitEmail, submitPassword, backToEmail, signOut };
}
