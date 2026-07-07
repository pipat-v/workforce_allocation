export type LoginSession = {
  username: string;
  position: string | null;
  menuAccess: string;
  loginAt: number;
  expiresAt: number;
};

const SESSION_KEY = "was_login_session";
export const SESSION_TTL_MS = 10 * 60 * 60 * 1000; // 10 ชั่วโมง

export function getSession(): LoginSession | null {
  if (typeof window === "undefined") return null;
  const raw = window.localStorage.getItem(SESSION_KEY);
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw) as LoginSession;
    if (!parsed.expiresAt || parsed.expiresAt <= Date.now()) {
      window.localStorage.removeItem(SESSION_KEY);
      return null;
    }
    return parsed;
  } catch {
    window.localStorage.removeItem(SESSION_KEY);
    return null;
  }
}

export function saveSession(user: {
  username: string;
  position: string | null;
  menu_access: string;
}): LoginSession {
  const now = Date.now();
  const session: LoginSession = {
    username: user.username,
    position: user.position,
    menuAccess: user.menu_access,
    loginAt: now,
    expiresAt: now + SESSION_TTL_MS,
  };
  window.localStorage.setItem(SESSION_KEY, JSON.stringify(session));
  return session;
}

export function clearSession() {
  if (typeof window === "undefined") return;
  window.localStorage.removeItem(SESSION_KEY);
}

// menuAccess is either "All" or a comma/space-separated list of menu numbers
// (e.g. "0,1,4"); decimal sub-menu numbers (e.g. "4.2") still grant the
// parent tab since gating only happens at the top-level tab.
export function hasMenuAccess(session: LoginSession | null, menuNo: number): boolean {
  if (!session) return false;
  const value = session.menuAccess?.trim();
  if (!value) return false;
  if (value.toLowerCase() === "all") return true;
  const allowed = value
    .split(/[,\s]+/)
    .map((part) => part.trim())
    .filter((part) => part !== "")
    .map((part) => Number(part))
    .filter((n) => !Number.isNaN(n));
  return allowed.some((n) => Math.trunc(n) === menuNo);
}
