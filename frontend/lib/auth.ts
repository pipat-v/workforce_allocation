export type LoginSession = {
  username: string;
  position: string | null;
  menuAccess: string;
  menuViewAccess: string;
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
  menu_view_access?: string;
}): LoginSession {
  const now = Date.now();
  const session: LoginSession = {
    username: user.username,
    position: user.position,
    menuAccess: user.menu_access,
    menuViewAccess: user.menu_view_access ?? "All",
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

// value is either "All" or a comma/space-separated list of menu numbers
// (e.g. "0,1,4"); decimal sub-menu numbers (e.g. "4.2") still grant the
// parent tab since gating only happens at the top-level tab.
export function isAllMenuAccess(value: string | undefined): boolean {
  return (value?.trim().toLowerCase() ?? "") === "all";
}

export function parseMenuNumbers(value: string | undefined): number[] {
  const trimmed = value?.trim();
  if (!trimmed || trimmed.toLowerCase() === "all") return [];
  return trimmed
    .split(/[,\s]+/)
    .map((part) => part.trim())
    .filter((part) => part !== "")
    .map((part) => Number(part))
    .filter((n) => !Number.isNaN(n));
}

function parseMenuList(value: string | undefined, menuNo: number): boolean {
  if (isAllMenuAccess(value)) return true;
  return parseMenuNumbers(value).some((n) => Math.trunc(n) === menuNo);
}

export function hasMenuAccess(session: LoginSession | null, menuNo: number): boolean {
  if (!session) return false;
  return parseMenuList(session.menuAccess, menuNo);
}

// Not enforced anywhere yet — every page stays publicly viewable regardless of
// this value. Exists so the Setting page's per-user View column has a real
// permission behind it, ready for future view-gating.
export function hasMenuViewAccess(session: LoginSession | null, menuNo: number): boolean {
  if (!session) return false;
  return parseMenuList(session.menuViewAccess, menuNo);
}

// The Setting page (managing users/passwords) is restricted to these
// positions specifically, not to whoever's menu_access happens to cover it.
const SETTING_ADMIN_POSITIONS = ["hr", "เถ้าแก่"];

export function canManageSettingUsers(session: LoginSession | null): boolean {
  if (!session) return false;
  const position = (session.position ?? "").trim().toLowerCase();
  return SETTING_ADMIN_POSITIONS.includes(position);
}
