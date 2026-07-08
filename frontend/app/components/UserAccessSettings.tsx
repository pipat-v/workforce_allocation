"use client";

import { useEffect, useRef, useState } from "react";
import type { FormEvent } from "react";
import { ChevronDown, Download, FileSpreadsheet, KeyRound, Plus, RotateCw, ShieldCheck, Trash2, UploadCloud, UserPlus, Users } from "lucide-react";
import * as XLSX from "xlsx";
import { supabase } from "@/lib/supabase";
import { canManageSettingUsers, isAllMenuAccess, parseMenuNumbers, type LoginSession } from "@/lib/auth";
import LoginGate from "./LoginGate";

type ParsedUser = {
  position: string;
  username: string;
  password: string;
  menu_access: string;
  menu_view_access: string;
};
type ParsedMenu = { menu_no: number; menu_name: string };

type LoginUserRow = ParsedUser & { id: number };
type LoginMenuRow = ParsedMenu & { id: number };

function menuAccessBadgeInfo(value: string, totalMenus: number): { label: string; variant: "all" | "partial" | "none" } {
  if (isAllMenuAccess(value)) return { label: "ทั้งหมด", variant: "all" };
  const count = parseMenuNumbers(value).length;
  if (count === 0) return { label: "ไม่มีสิทธิ์", variant: "none" };
  return { label: totalMenus > 0 ? `${count}/${totalMenus} ` + "เมนู" : `${count} เมนู`, variant: "partial" };
}

function menuNameListForAccess(value: string, menus: LoginMenuRow[]): string[] {
  const nameByNo = new Map(menus.map((m) => [m.menu_no, m.menu_name]));
  return parseMenuNumbers(value).map((n) => nameByNo.get(n) ?? String(n));
}

function AccessBadge({ value, menus }: { value: string; menus: LoginMenuRow[] }) {
  const { label, variant } = menuAccessBadgeInfo(value, menus.length);
  return (
    <div className="access-badge-cell">
      <span className={`access-badge ${variant}`}>{label}</span>
      {variant === "partial" ? (
        <div className="access-menu-taglist">
          {menuNameListForAccess(value, menus).map((name, idx) => (
            <span key={`${name}-${idx}`} className="access-menu-tag">
              {name}
            </span>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function normalizeHeader(h: string) {
  return h.trim().toLowerCase();
}

function findValue(row: Record<string, unknown>, key: string): unknown {
  const found = Object.keys(row).find((k) => normalizeHeader(k) === key);
  return found ? row[found] : "";
}

function parseUserFile(buffer: ArrayBuffer): { users: ParsedUser[]; menus: ParsedMenu[] } {
  const workbook = XLSX.read(buffer, { type: "array" });
  const mainSheetName =
    workbook.SheetNames.find((n) => normalizeHeader(n) === "main") ?? workbook.SheetNames[0];
  const detailSheetName = workbook.SheetNames.find((n) => normalizeHeader(n) === "detail");

  const mainRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[mainSheetName], {
    defval: "",
  });
  const usersByUsername = new Map<string, ParsedUser>();
  for (const row of mainRows) {
    const menuAccess = String(findValue(row, "menu") ?? "").trim() || "All";
    const parsedRow = {
      position: String(findValue(row, "ตำแหน่ง") ?? "").trim(),
      username: String(findValue(row, "user") ?? "").trim(),
      password: String(findValue(row, "password") ?? "").trim(),
      menu_access: menuAccess,
      // The Excel template only has one Menu column — bulk import grants the
      // same access for View as for Edit. Fine-tuning View separately from
      // Edit for one person is done afterward via the per-user matrix.
      menu_view_access: menuAccess,
    };
    if (!parsedRow.username || !parsedRow.password) continue;
    // last row for a given username wins, so the sheet's unique(username)
    // constraint can't reject the insert after the old rows are already deleted
    usersByUsername.set(parsedRow.username, parsedRow);
  }
  const users = Array.from(usersByUsername.values());

  const menusByNo = new Map<number, ParsedMenu>();
  if (detailSheetName) {
    const detailRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
      workbook.Sheets[detailSheetName],
      { defval: "" },
    );
    for (const row of detailRows) {
      const menuName = String(findValue(row, "เมนู") ?? "").trim();
      const menuNo = Number(findValue(row, "หมายเลข"));
      if (!menuName || Number.isNaN(menuNo)) continue;
      menusByNo.set(menuNo, { menu_no: menuNo, menu_name: menuName });
    }
  }
  const menus = Array.from(menusByNo.values());

  return { users, menus };
}

// Shared Edit/View-per-menu checkbox grid state, used both by the per-user
// permissions modal and the add-user form. Each caller gets its own
// independent instance (call the hook once per usage site). Every menu is
// always individually selectable — "All" is just a derived shortcut (checked
// when every menu is checked), not a separate mode that hides the grid.
function useMenuAccessGrid(menus: LoginMenuRow[]) {
  const [editSet, setEditSet] = useState<Set<number>>(new Set());
  const [viewSet, setViewSet] = useState<Set<number>>(new Set());
  const didAutoInit = useRef(false);

  // currentMenus loads asynchronously after mount, so the grid starts empty —
  // once the menu list arrives, default a still-untouched grid to fully
  // checked (matches the "All" default new users have always gotten).
  useEffect(() => {
    if (!didAutoInit.current && menus.length > 0) {
      didAutoInit.current = true;
      setEditSet(new Set(menus.map((m) => m.menu_no)));
      setViewSet(new Set(menus.map((m) => m.menu_no)));
    }
  }, [menus]);

  function setFromAccess(menuAccess: string, menuViewAccess: string) {
    setEditSet(
      new Set(isAllMenuAccess(menuAccess) ? menus.map((m) => m.menu_no) : parseMenuNumbers(menuAccess)),
    );
    setViewSet(
      new Set(isAllMenuAccess(menuViewAccess) ? menus.map((m) => m.menu_no) : parseMenuNumbers(menuViewAccess)),
    );
  }

  function reset() {
    setEditSet(new Set(menus.map((m) => m.menu_no)));
    setViewSet(new Set(menus.map((m) => m.menu_no)));
  }

  function toggleEdit(menuNo: number) {
    setEditSet((prev) => {
      const next = new Set(prev);
      if (next.has(menuNo)) next.delete(menuNo);
      else next.add(menuNo);
      return next;
    });
  }
  function toggleView(menuNo: number) {
    setViewSet((prev) => {
      const next = new Set(prev);
      if (next.has(menuNo)) next.delete(menuNo);
      else next.add(menuNo);
      return next;
    });
  }
  function toggleAllEdit() {
    setEditSet((prev) => (prev.size >= menus.length ? new Set() : new Set(menus.map((m) => m.menu_no))));
  }
  function toggleAllView() {
    setViewSet((prev) => (prev.size >= menus.length ? new Set() : new Set(menus.map((m) => m.menu_no))));
  }

  // When menus hasn't loaded yet, treat "all" as vacuously true (nothing to
  // check yet) rather than false — otherwise toStrings() below would save an
  // empty "no access" grant if a submit races ahead of the menu list fetch.
  const allEdit = menus.length === 0 || editSet.size >= menus.length;
  const allView = menus.length === 0 || viewSet.size >= menus.length;

  function toStrings(): { menu_access: string; menu_view_access: string } {
    return {
      menu_access: allEdit ? "All" : Array.from(editSet).sort((a, b) => a - b).join(","),
      menu_view_access: allView ? "All" : Array.from(viewSet).sort((a, b) => a - b).join(","),
    };
  }

  return { allEdit, allView, editSet, viewSet, setFromAccess, reset, toggleAllEdit, toggleAllView, toggleEdit, toggleView, toStrings };
}

function MenuAccessGrid({ menus, grid }: { menus: LoginMenuRow[]; grid: ReturnType<typeof useMenuAccessGrid> }) {
  return (
    <div className="settings-table-wrap menu-access-grid">
      <table className="table compact-table">
        <thead>
          <tr>
            <th>เมนู</th>
            <th style={{ textAlign: "center" }}>
              <label className="menu-access-all-toggle">
                <input type="checkbox" checked={grid.allEdit} onChange={grid.toggleAllEdit} />
                แก้ไข (Edit) ทั้งหมด
              </label>
            </th>
            <th style={{ textAlign: "center" }}>
              <label className="menu-access-all-toggle">
                <input type="checkbox" checked={grid.allView} onChange={grid.toggleAllView} />
                ดู (View) ทั้งหมด
              </label>
            </th>
          </tr>
        </thead>
        <tbody>
          {menus.map((m) => (
            <tr key={m.menu_no} className={Number.isInteger(m.menu_no) ? undefined : "menu-access-submenu"}>
              <td>{m.menu_name}</td>
              <td style={{ textAlign: "center" }}>
                <input
                  type="checkbox"
                  checked={grid.editSet.has(m.menu_no)}
                  onChange={() => grid.toggleEdit(m.menu_no)}
                />
              </td>
              <td style={{ textAlign: "center" }}>
                <input
                  type="checkbox"
                  checked={grid.viewSet.has(m.menu_no)}
                  onChange={() => grid.toggleView(m.menu_no)}
                />
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// Compact preview shown instead of the full 19-row grid until the user asks
// to fine-tune — most adds/edits just want the default (or current) access
// as-is, so the big checkbox table shouldn't be forced on every open.
function MenuAccessSummary({
  grid,
  menus,
  onExpand,
}: {
  grid: ReturnType<typeof useMenuAccessGrid>;
  menus: LoginMenuRow[];
  onExpand: () => void;
}) {
  const { menu_access, menu_view_access } = grid.toStrings();
  return (
    <div className="access-quick-summary">
      <div className="access-quick-summary-badges">
        <span className="access-quick-summary-label">แก้ไข</span>
        <AccessBadge value={menu_access} menus={menus} />
        <span className="access-quick-summary-label">ดู</span>
        <AccessBadge value={menu_view_access} menus={menus} />
      </div>
      <button type="button" className="link-button small" onClick={onExpand}>
        ปรับแต่งรายเมนู
      </button>
    </div>
  );
}

export default function UserAccessSettings({
  session,
  onLoginSuccess,
}: {
  session: LoginSession | null;
  onLoginSuccess: (session: LoginSession) => void;
}) {
  const [currentUsers, setCurrentUsers] = useState<LoginUserRow[]>([]);
  const [currentMenus, setCurrentMenus] = useState<LoginMenuRow[]>([]);
  const [currentUsersCollapsed, setCurrentUsersCollapsed] = useState(false);
  const [loadError, setLoadError] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [parsed, setParsed] = useState<{ users: ParsedUser[]; menus: ParsedMenu[] } | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [passwordEditRow, setPasswordEditRow] = useState<LoginUserRow | null>(null);
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [isChangingPassword, setIsChangingPassword] = useState(false);

  const [permsEditRow, setPermsEditRow] = useState<LoginUserRow | null>(null);
  const [permsError, setPermsError] = useState("");
  const [isSavingPerms, setIsSavingPerms] = useState(false);
  const [showPermsGrid, setShowPermsGrid] = useState(false);
  const permsGrid = useMenuAccessGrid(currentMenus);

  const [showAddUser, setShowAddUser] = useState(false);
  const [showAddGrid, setShowAddGrid] = useState(false);
  const [newPosition, setNewPosition] = useState("");
  const [newUsername, setNewUsername] = useState("");
  const [newUserPassword, setNewUserPassword] = useState("");
  const [addUserError, setAddUserError] = useState("");
  const [isAddingUser, setIsAddingUser] = useState(false);
  const addGrid = useMenuAccessGrid(currentMenus);

  // Only HR / เถ้าแก่ can manage users/passwords here — having menu_access
  // to menu 8 is not enough (e.g. an "All"-access account with another position).
  const canEdit = canManageSettingUsers(session);

  async function loadCurrent() {
    setIsLoading(true);
    const [{ data: users, error: usersError }, { data: menus, error: menusError }] = await Promise.all([
      supabase
        .from("login_users")
        .select("id, position, username, menu_access, menu_view_access")
        .order("id"),
      supabase.from("login_menus").select("id, menu_no, menu_name").order("menu_no"),
    ]);
    const failures: string[] = [];
    if (usersError) failures.push(`ผู้ใช้: ${usersError.message}`);
    else if (users) setCurrentUsers(users as LoginUserRow[]);
    if (menusError) failures.push(`เมนู: ${menusError.message}`);
    else if (menus) setCurrentMenus(menus as LoginMenuRow[]);
    setLoadError(failures.join(" / "));
    setIsLoading(false);
  }

  useEffect(() => {
    void loadCurrent();
  }, []);

  function downloadUserTemplate() {
    const headers = ["ตำแหน่ง", "User", "Password", "Menu"];
    const examples = [
      ["HR", "somchai", "1234", "All"],
      ["พนักงาน", "somsri", "1234", "0,1,4"],
    ];
    const ws = XLSX.utils.aoa_to_sheet([headers, ...examples]);
    ws["!cols"] = [16, 14, 12, 20].map((w) => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Main");

    const detailRows = currentMenus.length > 0
      ? currentMenus.map((m) => [m.menu_no, m.menu_name])
      : [[0, "Dashboard"]];
    const detailWs = XLSX.utils.aoa_to_sheet([["หมายเลข", "เมนู"], ...detailRows]);
    detailWs["!cols"] = [10, 22].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, detailWs, "Detail");

    XLSX.writeFile(wb, "User ระบบคน-template.xlsx");
  }

  async function handleFileChange(selected: File | null) {
    setFile(selected);
    setParsed(null);
    setError("");
    setMessage("");
    if (!selected) return;
    try {
      const buffer = await selected.arrayBuffer();
      const result = parseUserFile(buffer);
      if (!result.users.length) {
        setError('ไม่พบข้อมูลผู้ใช้ในชีท "Main" — ตรวจสอบหัวคอลัมน์ ตำแหน่ง/User/Password/Menu');
        return;
      }
      setParsed(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "อ่านไฟล์ไม่สำเร็จ");
    }
  }

  async function handleSave() {
    if (!parsed) return;
    setIsSaving(true);
    setError("");
    setMessage("");
    try {
      // The uploaded sheet only has one Menu column, so a fresh parse always
      // sets menu_view_access = menu_access. For usernames that already exist,
      // preserve whatever View access they currently have in the DB (which may
      // have been fine-tuned separately via the per-user matrix) instead of
      // clobbering it back to match Edit on every re-upload.
      const existingViewAccessByUsername = new Map(currentUsers.map((u) => [u.username, u.menu_view_access]));
      const usersToInsert = parsed.users.map((u) => {
        const existingView = existingViewAccessByUsername.get(u.username);
        return existingView !== undefined ? { ...u, menu_view_access: existingView } : u;
      });

      const { error: delUsersErr } = await supabase.from("login_users").delete().gte("id", 0);
      if (delUsersErr) throw new Error(delUsersErr.message);
      if (usersToInsert.length) {
        const { error: insUsersErr } = await supabase.from("login_users").insert(usersToInsert);
        if (insUsersErr) throw new Error(insUsersErr.message);
      }

      // Only touch login_menus when the workbook actually had a readable "Detail" sheet —
      // otherwise a file missing/misnamed that sheet would silently wipe the menu reference list.
      if (parsed.menus.length) {
        const { error: delMenusErr } = await supabase.from("login_menus").delete().gte("id", 0);
        if (delMenusErr) throw new Error(delMenusErr.message);
        const { error: insMenusErr } = await supabase.from("login_menus").insert(parsed.menus);
        if (insMenusErr) throw new Error(insMenusErr.message);
      }

      setMessage(`บันทึกสำเร็จ: ${parsed.users.length} ผู้ใช้, ${parsed.menus.length} เมนู`);
      setFile(null);
      setParsed(null);
      await loadCurrent();
    } catch (err) {
      setError(err instanceof Error ? err.message : "บันทึกไม่สำเร็จ");
    } finally {
      setIsSaving(false);
    }
  }

  async function handleDeleteUser(id: number) {
    if (!window.confirm("ต้องการลบผู้ใช้นี้?")) return;
    const { error: delErr } = await supabase.from("login_users").delete().eq("id", id);
    if (delErr) {
      alert(`ลบไม่สำเร็จ: ${delErr.message}`);
      return;
    }
    await loadCurrent();
  }

  function openPasswordEdit(row: LoginUserRow) {
    setPasswordEditRow(row);
    setNewPassword("");
    setConfirmPassword("");
    setPasswordError("");
  }

  async function handleChangePassword() {
    if (!passwordEditRow) return;
    if (!newPassword.trim()) {
      setPasswordError("กรุณากรอกรหัสผ่านใหม่");
      return;
    }
    if (newPassword !== confirmPassword) {
      setPasswordError("รหัสผ่านทั้งสองช่องไม่ตรงกัน");
      return;
    }
    setIsChangingPassword(true);
    setPasswordError("");
    const { error: updateErr } = await supabase
      .from("login_users")
      .update({ password: newPassword.trim(), updated_at: new Date().toISOString() })
      .eq("id", passwordEditRow.id);
    setIsChangingPassword(false);
    if (updateErr) {
      setPasswordError(updateErr.message);
      return;
    }
    setPasswordEditRow(null);
    setMessage(`เปลี่ยนรหัสผ่านของ ${passwordEditRow.username} แล้ว`);
  }

  function openPermsEdit(row: LoginUserRow) {
    setPermsEditRow(row);
    setPermsError("");
    setShowPermsGrid(false);
    permsGrid.setFromAccess(row.menu_access, row.menu_view_access);
  }

  async function handleSavePerms() {
    if (!permsEditRow) return;
    setIsSavingPerms(true);
    setPermsError("");
    const { menu_access, menu_view_access } = permsGrid.toStrings();
    const { error: updErr } = await supabase
      .from("login_users")
      .update({ menu_access, menu_view_access, updated_at: new Date().toISOString() })
      .eq("id", permsEditRow.id);
    setIsSavingPerms(false);
    if (updErr) {
      setPermsError(updErr.message);
      return;
    }
    setMessage(`อัพเดทสิทธิ์ของ ${permsEditRow.username} แล้ว`);
    setPermsEditRow(null);
    await loadCurrent();
  }

  async function handleAddUser(e: FormEvent) {
    e.preventDefault();
    if (!newUsername.trim() || !newUserPassword) {
      setAddUserError("กรุณากรอก User และ Password");
      return;
    }
    setIsAddingUser(true);
    setAddUserError("");
    const { menu_access, menu_view_access } = addGrid.toStrings();
    const { error: insErr } = await supabase.from("login_users").insert([
      {
        position: newPosition.trim() || null,
        username: newUsername.trim(),
        password: newUserPassword,
        menu_access,
        menu_view_access,
      },
    ]);
    setIsAddingUser(false);
    if (insErr) {
      setAddUserError(insErr.code === "23505" ? "มีชื่อผู้ใช้นี้อยู่แล้ว" : insErr.message);
      return;
    }
    setNewPosition("");
    setNewUsername("");
    setNewUserPassword("");
    addGrid.reset();
    setShowAddGrid(false);
    setShowAddUser(false);
    setMessage(`เพิ่มผู้ใช้ ${newUsername.trim()} แล้ว`);
    await loadCurrent();
  }

  return (
    <section className="panel settings-page">
      <div className="settings-page-header">
        <div className="settings-current-header">
          <h3>ผู้ใช้ระบบ & สิทธิ์การเข้าเมนู</h3>
          {canEdit ? (
            <button className="secondary-button small" type="button" onClick={() => void loadCurrent()} disabled={isLoading}>
              <RotateCw size={14} />
              รีเฟรช
            </button>
          ) : null}
        </div>
        <p>ดูรายชื่อผู้ใช้และแก้ไข/เพิ่ม/ลบผู้ใช้ได้เฉพาะบัญชีตำแหน่ง &quot;HR&quot;, &quot;เถ้าแก่&quot; หรือ &quot;ผู้จัดการ&quot; เท่านั้น</p>
      </div>

      {loadError ? <p className="login-gate-error">โหลดข้อมูลไม่สำเร็จ: {loadError}</p> : null}

      {!canEdit ? (
        session ? (
          <p className="login-gate-error">บัญชี {session.username} ไม่มีสิทธิ์เข้าถึงส่วนนี้ (เฉพาะ HR, เถ้าแก่ หรือ ผู้จัดการ เท่านั้น)</p>
        ) : (
          <LoginGate menuLabel="จัดการผู้ใช้ระบบ" onSuccess={onLoginSuccess} />
        )
      ) : (
        <>
          <div className="settings-section">
            <div className="settings-section-header">
              <h4>
                <UserPlus size={16} />
                เพิ่มผู้ใช้ใหม่
              </h4>
              <button className="secondary-button small" type="button" onClick={() => setShowAddUser((v) => !v)}>
                <Plus size={14} />
                {showAddUser ? "ยกเลิก" : "เพิ่มผู้ใช้"}
              </button>
            </div>

            {showAddUser ? (
              <form className="login-gate-form" onSubmit={(e) => void handleAddUser(e)}>
                <div className="login-gate-field-row">
                  <label className="login-gate-field">
                    <span>ตำแหน่ง</span>
                    <input type="text" value={newPosition} onChange={(e) => setNewPosition(e.target.value)} />
                  </label>
                  <label className="login-gate-field">
                    <span>User</span>
                    <input type="text" value={newUsername} onChange={(e) => setNewUsername(e.target.value)} autoComplete="off" />
                  </label>
                  <label className="login-gate-field">
                    <span>Password</span>
                    <input
                      type="password"
                      value={newUserPassword}
                      onChange={(e) => setNewUserPassword(e.target.value)}
                      autoComplete="new-password"
                    />
                  </label>
                </div>
                {showAddGrid ? (
                  <MenuAccessGrid menus={currentMenus} grid={addGrid} />
                ) : (
                  <MenuAccessSummary grid={addGrid} menus={currentMenus} onExpand={() => setShowAddGrid(true)} />
                )}
                {addUserError ? <p className="login-gate-error">{addUserError}</p> : null}
                <button className="primary-button" type="submit" disabled={isAddingUser}>
                  {isAddingUser ? "กำลังบันทึก..." : "เพิ่มผู้ใช้"}
                </button>
              </form>
            ) : null}
          </div>

          <div className="settings-section">
            <div className="settings-section-header">
              <h4>
                <FileSpreadsheet size={16} />
                อัปโหลดไฟล์ Excel (เพิ่มหลายคนพร้อมกัน)
              </h4>
              <button className="secondary-button small" type="button" onClick={downloadUserTemplate}>
                <Download size={14} />
                ดาวน์โหลด Template
              </button>
            </div>
            <p className="settings-upload-hint">
              ไฟล์ &quot;User ระบบคน.xlsx&quot; — ชีท Main: ตำแหน่ง/User/Password/Menu, ชีท Detail: หมายเลข/เมนู — การบันทึกจะแทนที่ข้อมูลผู้ใช้ทั้งหมดด้วยไฟล์ที่อัพโหลด
              ผู้ใช้ใหม่จะได้สิทธิ์ Edit กับ View เท่ากันตามคอลัมน์ Menu ส่วนผู้ใช้ที่มีอยู่แล้วจะคงสิทธิ์ View เดิมไว้ (ไม่ถูกทับด้วยค่าจากไฟล์)
            </p>
            <label className={`ts-dropzone compact-dropzone${file ? " has-file" : ""}`}>
              <UploadCloud size={24} />
              {file ? (
                <>
                  <strong>{file.name}</strong>
                  <span>{(file.size / 1024).toFixed(0)} KB · คลิกเพื่อเปลี่ยนไฟล์</span>
                </>
              ) : (
                <>
                  <strong>ลากไฟล์มาวางที่นี่ หรือ คลิกเพื่อเลือกไฟล์</strong>
                  <span>รองรับ .xlsx, .xls</span>
                </>
              )}
              <input
                type="file"
                accept=".xlsx,.xls"
                style={{ display: "none" }}
                onChange={(e) => void handleFileChange(e.target.files?.[0] ?? null)}
              />
            </label>

            {error ? <p className="login-gate-error">{error}</p> : null}
            {message ? <p className="settings-success">{message}</p> : null}

            {parsed ? (
              <div className="settings-preview">
                <div className="settings-preview-header">
                  <FileSpreadsheet size={16} />
                  <span>พบผู้ใช้ {parsed.users.length} รายการ, เมนู {parsed.menus.length} รายการ — ตรวจสอบก่อนบันทึก</span>
                </div>
                <div className="settings-table-wrap">
                  <table className="table data-table">
                    <thead>
                      <tr>
                        <th>ตำแหน่ง</th>
                        <th>User</th>
                        <th>Menu (Edit=View)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {parsed.users.map((row, idx) => (
                        <tr key={`${row.username}-${idx}`}>
                          <td>{row.position || "-"}</td>
                          <td>{row.username}</td>
                          <td>{row.menu_access}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <button className="primary-button" type="button" disabled={isSaving} onClick={() => void handleSave()}>
                  {isSaving ? "กำลังบันทึก..." : "บันทึก (แทนที่ข้อมูลผู้ใช้ทั้งหมด)"}
                </button>
              </div>
            ) : null}
          </div>

          <div className="settings-section settings-current">
            <div className="settings-section-header panel-collapse-trigger" onClick={() => setCurrentUsersCollapsed((c) => !c)}>
              <h4>
                <Users size={16} />
                ผู้ใช้ที่ใช้งานอยู่ในระบบ
              </h4>
              <ChevronDown
                size={16}
                className="panel-collapse-chevron"
                style={{ transform: currentUsersCollapsed ? "rotate(-90deg)" : "rotate(0deg)" }}
              />
            </div>
            {!currentUsersCollapsed ? (
              <div className="settings-table-wrap">
                <table className="table">
                  <thead>
                    <tr>
                      <th>ตำแหน่ง</th>
                      <th>User</th>
                      <th>Edit</th>
                      <th>View</th>
                      <th></th>
                    </tr>
                  </thead>
                  <tbody>
                    {currentUsers.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="empty-copy">
                          {isLoading ? "กำลังโหลด..." : "ยังไม่มีผู้ใช้"}
                        </td>
                      </tr>
                    ) : (
                      currentUsers.map((row) => (
                        <tr key={row.id}>
                          <td>{row.position || "-"}</td>
                          <td>{row.username}</td>
                          <td><AccessBadge value={row.menu_access} menus={currentMenus} /></td>
                          <td><AccessBadge value={row.menu_view_access} menus={currentMenus} /></td>
                          <td>
                            <div style={{ display: "flex", gap: 4 }}>
                              <button className="icon-button" type="button" title="สิทธิ์ Edit/View" onClick={() => openPermsEdit(row)}>
                                <ShieldCheck size={15} />
                              </button>
                              <button className="icon-button" type="button" title="เปลี่ยนรหัสผ่าน" onClick={() => openPasswordEdit(row)}>
                                <KeyRound size={15} />
                              </button>
                              <button className="icon-button danger" type="button" title="ลบ" onClick={() => void handleDeleteUser(row.id)}>
                                <Trash2 size={15} />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            ) : null}
          </div>
        </>
      )}

      {passwordEditRow ? (
        <div className="modal-overlay" onClick={() => setPasswordEditRow(null)}>
          <section className="panel login-gate" onClick={(e) => e.stopPropagation()}>
            <div className="login-gate-icon">
              <KeyRound size={22} />
            </div>
            <h3>เปลี่ยนรหัสผ่านของ {passwordEditRow.username}</h3>
            <form
              className="login-gate-form"
              onSubmit={(e) => {
                e.preventDefault();
                void handleChangePassword();
              }}
            >
              <label className="login-gate-field">
                <span>รหัสผ่านใหม่</span>
                <input
                  type="password"
                  value={newPassword}
                  onChange={(e) => setNewPassword(e.target.value)}
                  autoFocus
                />
              </label>
              <label className="login-gate-field">
                <span>ยืนยันรหัสผ่านใหม่</span>
                <input
                  type="password"
                  value={confirmPassword}
                  onChange={(e) => setConfirmPassword(e.target.value)}
                />
              </label>
              {passwordError ? <p className="login-gate-error">{passwordError}</p> : null}
              <button className="primary-button" type="submit" disabled={isChangingPassword}>
                {isChangingPassword ? "กำลังบันทึก..." : "บันทึกรหัสผ่านใหม่"}
              </button>
            </form>
          </section>
        </div>
      ) : null}

      {permsEditRow ? (
        <div className="modal-overlay" onClick={() => setPermsEditRow(null)}>
          <section className="panel login-gate settings-perms-modal" onClick={(e) => e.stopPropagation()}>
            <div className="login-gate-icon">
              <ShieldCheck size={22} />
            </div>
            <h3>สิทธิ์ Edit/View ของ {permsEditRow.username}</h3>
            {showPermsGrid ? (
              <MenuAccessGrid menus={currentMenus} grid={permsGrid} />
            ) : (
              <MenuAccessSummary grid={permsGrid} menus={currentMenus} onExpand={() => setShowPermsGrid(true)} />
            )}
            {permsError ? <p className="login-gate-error">{permsError}</p> : null}
            <button
              className="primary-button"
              type="button"
              disabled={isSavingPerms}
              onClick={() => void handleSavePerms()}
              style={{ marginTop: 12 }}
            >
              {isSavingPerms ? "กำลังบันทึก..." : "บันทึกสิทธิ์"}
            </button>
          </section>
        </div>
      ) : null}
    </section>
  );
}
