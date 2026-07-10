"use client";

import { useState } from "react";
import type { FormEvent } from "react";
import { Lock } from "lucide-react";
import { supabase } from "@/lib/supabase";
import { saveSession, type LoginSession } from "@/lib/auth";

const REGISTER_POSITION_OPTIONS = ["หมวกส้ม", "เจ้าหน้าที่ฝ่าย(Staff)", "ผู้บริหาร"];

export default function LoginGate({
  menuLabel,
  onSuccess,
  variant = "modal",
}: {
  menuLabel: string;
  onSuccess: (session: LoginSession) => void;
  variant?: "modal" | "page";
}) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const [showRegister, setShowRegister] = useState(false);
  const [regPosition, setRegPosition] = useState("");
  const [regUsername, setRegUsername] = useState("");
  const [regPassword, setRegPassword] = useState("");
  const [regConfirmPassword, setRegConfirmPassword] = useState("");
  const [regError, setRegError] = useState("");
  const [regMessage, setRegMessage] = useState("");
  const [regLoading, setRegLoading] = useState(false);

  function openRegister() {
    setShowRegister(true);
    setRegError("");
    setRegMessage("");
    setRegPosition("");
    setRegUsername("");
    setRegPassword("");
    setRegConfirmPassword("");
  }

  async function handleRegisterSubmit(e: FormEvent) {
    e.preventDefault();
    if (!regPosition) {
      setRegError("กรุณาเลือกตำแหน่ง");
      return;
    }
    if (!regUsername.trim() || !regPassword) {
      setRegError("กรุณากรอก User และ Password");
      return;
    }
    if (regPassword !== regConfirmPassword) {
      setRegError("รหัสผ่านทั้งสองช่องไม่ตรงกัน");
      return;
    }
    setRegLoading(true);
    setRegError("");
    try {
      const { error: existingErr, data: existing } = await supabase
        .from("login_users")
        .select("username")
        .eq("username", regUsername.trim())
        .maybeSingle();
      if (existingErr) throw new Error(existingErr.message);
      if (existing) throw new Error("มีชื่อผู้ใช้นี้อยู่แล้ว");

      const { error: insErr } = await supabase.from("registration_requests").insert([
        {
          position: regPosition.trim() || null,
          username: regUsername.trim(),
          password: regPassword,
        },
      ]);
      if (insErr) throw new Error(insErr.message);
      setRegMessage("ส่งคำขอลงทะเบียนแล้ว รอการอนุมัติจาก HR, เถ้าแก่ หรือ ผู้จัดการ");
      setRegPosition("");
      setRegUsername("");
      setRegPassword("");
      setRegConfirmPassword("");
    } catch (err) {
      setRegError(err instanceof Error ? err.message : "ส่งคำขอไม่สำเร็จ");
    } finally {
      setRegLoading(false);
    }
  }

  async function handleSubmit(e: FormEvent) {
    e.preventDefault();
    if (!username.trim() || !password) {
      setError("กรุณากรอก User และ Password");
      return;
    }
    setLoading(true);
    setError("");
    try {
      const { data, error: dbError } = await supabase
        .from("login_users")
        .select("username, position, password, menu_access, menu_view_access")
        .eq("username", username.trim())
        .maybeSingle();
      if (dbError) throw new Error(dbError.message);
      if (!data || String(data.password) !== password) {
        setError("User หรือ Password ไม่ถูกต้อง");
        return;
      }
      const session = saveSession(data);
      onSuccess(session);
    } catch (err) {
      setError(err instanceof Error ? err.message : "เข้าสู่ระบบไม่สำเร็จ");
    } finally {
      setLoading(false);
    }
  }

  const fields = (
    <>
      <label className="login-gate-field">
        <span>User</span>
        <input
          type="text"
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          autoComplete="username"
          autoFocus
        />
      </label>
      <label className="login-gate-field">
        <span>Password</span>
        <input
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          autoComplete="current-password"
        />
      </label>
      {error ? <p className="login-gate-error">{error}</p> : null}
      <button className="primary-button" type="submit" disabled={loading}>
        {loading ? "กำลังตรวจสอบ..." : "เข้าสู่ระบบ"}
      </button>
    </>
  );

  if (variant === "page") {
    if (showRegister) {
      return (
        <div className="login-form-page">
          <h2>ลงทะเบียน</h2>
          <p className="login-form-page-desc">กรอกข้อมูลเพื่อขอสิทธิ์เข้าใช้งาน {menuLabel}</p>
          <form className="login-gate-form" onSubmit={(e) => void handleRegisterSubmit(e)}>
            <label className="login-gate-field">
              <span>ตำแหน่ง</span>
              <select
                value={regPosition}
                onChange={(e) => setRegPosition(e.target.value)}
                autoFocus
                required
              >
                <option value="">
                  — เลือกตำแหน่ง —
                </option>
                {REGISTER_POSITION_OPTIONS.map((position) => (
                  <option key={position} value={position}>{position}</option>
                ))}
              </select>
            </label>
            <label className="login-gate-field">
              <span>Username</span>
              <input
                type="text"
                value={regUsername}
                onChange={(e) => setRegUsername(e.target.value)}
                autoComplete="username"
                placeholder="เช่น somchai.jai"
              />
            </label>
            <label className="login-gate-field">
              <span>Password</span>
              <input
                type="password"
                value={regPassword}
                onChange={(e) => setRegPassword(e.target.value)}
                autoComplete="new-password"
              />
            </label>
            <label className="login-gate-field">
              <span>ยืนยัน Password</span>
              <input
                type="password"
                value={regConfirmPassword}
                onChange={(e) => setRegConfirmPassword(e.target.value)}
                autoComplete="new-password"
              />
            </label>
            {regError ? <p className="login-gate-error">{regError}</p> : null}
            {regMessage ? <p className="settings-success">{regMessage}</p> : null}
            <button className="primary-button" type="submit" disabled={regLoading}>
              {regLoading ? "กำลังส่งคำขอ..." : "ส่งคำขอลงทะเบียน"}
            </button>
            <button className="link-button" type="button" onClick={() => setShowRegister(false)}>
              กลับไปเข้าสู่ระบบ
            </button>
          </form>
        </div>
      );
    }
    return (
      <div className="login-form-page">
        <h2>เข้าสู่ระบบ</h2>
        <p className="login-form-page-desc">กรอกข้อมูลบัญชีเพื่อเข้าใช้งาน {menuLabel}</p>
        <form className="login-gate-form" onSubmit={handleSubmit}>
          {fields}
          <button className="link-button" type="button" onClick={openRegister}>
            ลงทะเบียน
          </button>
          <p className="login-form-page-note">
            *คำขอลงทะเบียน ต้องรอการอนุมัติจาก HR, เถ้าแก่ หรือ ผู้จัดการ
            <br />
            **หากลืมรหัสผ่าน ให้ลงทะเบียนใหม่ และแจ้งลบ Account เก่าออก
          </p>
        </form>
      </div>
    );
  }

  return (
    <section className="panel login-gate">
      <div className="login-gate-icon">
        <Lock size={22} />
      </div>
      <h3>เข้าสู่ระบบเพื่อใช้งาน &quot;{menuLabel}&quot;</h3>
      <p className="login-gate-desc">เมนูนี้ต้องเข้าสู่ระบบก่อนจึงจะเข้าใช้งานได้</p>
      <form className="login-gate-form" onSubmit={handleSubmit}>
        {fields}
      </form>
    </section>
  );
}
