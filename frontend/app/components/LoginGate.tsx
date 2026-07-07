"use client";

import { useState } from "react";
import type { FormEvent } from "react";
import { Lock } from "lucide-react";
import { supabase } from "@/lib/supabase";
import { saveSession, type LoginSession } from "@/lib/auth";

export default function LoginGate({
  menuLabel,
  onSuccess,
}: {
  menuLabel: string;
  onSuccess: (session: LoginSession) => void;
}) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

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
        .select("username, position, password, menu_access")
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

  return (
    <section className="panel login-gate">
      <div className="login-gate-icon">
        <Lock size={22} />
      </div>
      <h3>เข้าสู่ระบบเพื่อใช้งาน &quot;{menuLabel}&quot;</h3>
      <p className="login-gate-desc">เมนูนี้ต้องเข้าสู่ระบบก่อนจึงจะเข้าใช้งานได้</p>
      <form className="login-gate-form" onSubmit={handleSubmit}>
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
      </form>
    </section>
  );
}
