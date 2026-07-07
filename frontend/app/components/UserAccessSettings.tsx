"use client";

import { useEffect, useState } from "react";
import { FileSpreadsheet, KeyRound, RotateCw, Trash2, UploadCloud } from "lucide-react";
import * as XLSX from "xlsx";
import { supabase } from "@/lib/supabase";
import { hasMenuAccess, type LoginSession } from "@/lib/auth";
import LoginGate from "./LoginGate";

const SETTING_MENU_NO = 8;

type ParsedUser = { position: string; username: string; password: string; menu_access: string };
type ParsedMenu = { menu_no: number; menu_name: string };

type LoginUserRow = ParsedUser & { id: number };
type LoginMenuRow = ParsedMenu & { id: number };

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
    const parsedRow = {
      position: String(findValue(row, "ตำแหน่ง") ?? "").trim(),
      username: String(findValue(row, "user") ?? "").trim(),
      password: String(findValue(row, "password") ?? "").trim(),
      menu_access: String(findValue(row, "menu") ?? "").trim() || "All",
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

export default function UserAccessSettings({
  session,
  onLoginSuccess,
}: {
  session: LoginSession | null;
  onLoginSuccess: (session: LoginSession) => void;
}) {
  const [currentUsers, setCurrentUsers] = useState<LoginUserRow[]>([]);
  const [currentMenus, setCurrentMenus] = useState<LoginMenuRow[]>([]);
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
  const canEdit = !!session && hasMenuAccess(session, SETTING_MENU_NO);

  async function loadCurrent() {
    setIsLoading(true);
    const [{ data: users, error: usersError }, { data: menus, error: menusError }] = await Promise.all([
      supabase.from("login_users").select("id, position, username, menu_access").order("id"),
      supabase.from("login_menus").select("id, menu_no, menu_name").order("menu_no"),
    ]);
    if (!usersError && users) setCurrentUsers(users as LoginUserRow[]);
    if (!menusError && menus) setCurrentMenus(menus as LoginMenuRow[]);
    setIsLoading(false);
  }

  useEffect(() => {
    void loadCurrent();
  }, []);

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
      const { error: delUsersErr } = await supabase.from("login_users").delete().gte("id", 0);
      if (delUsersErr) throw new Error(delUsersErr.message);
      if (parsed.users.length) {
        const { error: insUsersErr } = await supabase.from("login_users").insert(parsed.users);
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

  return (
    <section className="panel settings-page">
      <div className="settings-page-header">
        <h3>ผู้ใช้ระบบ & สิทธิ์การเข้าเมนู</h3>
        <p>
          ดูรายชื่อผู้ใช้และเมนูได้โดยไม่ต้องเข้าสู่ระบบ แต่การอัพโหลด/แก้ไข/ลบ ต้องเข้าสู่ระบบก่อน
          ไฟล์ที่อัพโหลดคือ &quot;User ระบบคน.xlsx&quot; (ชีท Main: ตำแหน่ง/User/Password/Menu, ชีท Detail: หมายเลข/เมนู)
          การบันทึกจะแทนที่ข้อมูลผู้ใช้ทั้งหมดด้วยไฟล์ที่อัพโหลด
        </p>
      </div>

      {canEdit ? (
        <>
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
                      <th>Menu</th>
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
        </>
      ) : session ? (
        <p className="login-gate-error">บัญชี {session.username} ไม่มีสิทธิ์แก้ไขผู้ใช้ระบบ</p>
      ) : (
        <LoginGate menuLabel="แก้ไขผู้ใช้ระบบ" onSuccess={onLoginSuccess} />
      )}

      <div className="settings-current">
        <div className="settings-current-header">
          <h4>ผู้ใช้ที่ใช้งานอยู่ในระบบ</h4>
          <button className="secondary-button small" type="button" onClick={() => void loadCurrent()} disabled={isLoading}>
            <RotateCw size={14} />
            รีเฟรช
          </button>
        </div>
        <div className="settings-table-wrap">
          <table className="table data-table">
            <thead>
              <tr>
                <th>ตำแหน่ง</th>
                <th>User</th>
                <th>Menu</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {currentUsers.length === 0 ? (
                <tr>
                  <td colSpan={4} className="empty-copy">
                    {isLoading ? "กำลังโหลด..." : "ยังไม่มีผู้ใช้"}
                  </td>
                </tr>
              ) : (
                currentUsers.map((row) => (
                  <tr key={row.id}>
                    <td>{row.position || "-"}</td>
                    <td>{row.username}</td>
                    <td>{row.menu_access}</td>
                    <td>
                      {canEdit ? (
                        <div style={{ display: "flex", gap: 4 }}>
                          <button className="icon-button" type="button" title="เปลี่ยนรหัสผ่าน" onClick={() => openPasswordEdit(row)}>
                            <KeyRound size={15} />
                          </button>
                          <button className="icon-button danger" type="button" title="ลบ" onClick={() => void handleDeleteUser(row.id)}>
                            <Trash2 size={15} />
                          </button>
                        </div>
                      ) : null}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {currentMenus.length > 0 ? (
          <>
            <h4>รายชื่อเมนู (อ้างอิง)</h4>
            <div className="settings-table-wrap">
              <table className="table data-table compact-table">
                <thead>
                  <tr>
                    <th>หมายเลข</th>
                    <th>เมนู</th>
                  </tr>
                </thead>
                <tbody>
                  {currentMenus.map((row) => (
                    <tr key={row.id}>
                      <td>{row.menu_no}</td>
                      <td>{row.menu_name}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        ) : null}
      </div>

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
    </section>
  );
}
