"use client";

import { useEffect, useMemo, useState } from "react";
import { Database, UploadCloud, UsersRound } from "lucide-react";
import { supabase } from "@/lib/supabase";
import type { User } from "@supabase/supabase-js";

type AllocationRun = {
  id: string;
  target_date: string | null;
  status: string;
  scan_file_path: string | null;
  solver_status: string | null;
  created_at: string;
};

const requiredFiles = [
  { key: "scan", label: "Timestamp / Time Record" },
  { key: "master", label: "Master Employee" },
  { key: "manpower", label: "Manpower Plan" },
  { key: "skill", label: "Skill Matrix" },
] as const;

type FileKey = (typeof requiredFiles)[number]["key"];

export default function Home() {
  const [user, setUser] = useState<User | null>(null);
  const [email, setEmail] = useState("");
  const [files, setFiles] = useState<Record<FileKey, File | null>>({
    scan: null,
    master: null,
    manpower: null,
    skill: null,
  });
  const [runs, setRuns] = useState<AllocationRun[]>([]);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [isUploading, setIsUploading] = useState(false);

  const canUpload = useMemo(
    () => requiredFiles.every((item) => files[item.key]),
    [files],
  );

  useEffect(() => {
    supabase.auth.getUser().then(({ data }) => {
      setUser(data.user);
    });

    const { data: subscription } = supabase.auth.onAuthStateChange(
      (_event, session) => {
        setUser(session?.user ?? null);
      },
    );

    void loadRuns();

    return () => {
      subscription.subscription.unsubscribe();
    };
  }, []);

  async function sendMagicLink() {
    setError("");
    setMessage("");

    const { error: signInError } = await supabase.auth.signInWithOtp({
      email,
      options: {
        emailRedirectTo: window.location.origin,
      },
    });

    if (signInError) {
      setError(signInError.message);
      return;
    }

    setMessage("ส่งลิงก์เข้าใช้งานไปที่อีเมลแล้ว");
  }

  async function loadRuns() {
    const { data, error: loadError } = await supabase
      .from("allocation_runs")
      .select("id,target_date,status,scan_file_path,solver_status,created_at")
      .order("created_at", { ascending: false })
      .limit(20);

    if (loadError) {
      setError(loadError.message);
      return;
    }

    setRuns(data ?? []);
  }

  async function uploadFiles() {
    setError("");
    setMessage("");
    setIsUploading(true);

    if (!user) {
      setError("กรุณา login Supabase Auth ก่อน upload ไฟล์");
      setIsUploading(false);
      return;
    }

    const runId = crypto.randomUUID();
    const paths: Record<FileKey, string> = {
      scan: "",
      master: "",
      manpower: "",
      skill: "",
    };

    for (const item of requiredFiles) {
      const file = files[item.key];
      if (!file) continue;

      const path = `${user.id}/${runId}/${item.key}-${file.name}`;
      const { error: uploadError } = await supabase.storage
        .from("workforce-inputs")
        .upload(path, file, { upsert: true });

      if (uploadError) {
        setError(uploadError.message);
        setIsUploading(false);
        return;
      }

      paths[item.key] = path;
    }

    const { error: insertError } = await supabase.from("allocation_runs").insert({
      id: runId,
      owner_id: user.id,
      status: "uploaded",
      scan_file_path: paths.scan,
      master_file_path: paths.master,
      manpower_file_path: paths.manpower,
      skill_file_path: paths.skill,
    });

    if (insertError) {
      setError(insertError.message);
      setIsUploading(false);
      return;
    }

    setMessage("อัปโหลดแล้ว รอ worker ประมวลผล allocation run นี้");
    setIsUploading(false);
    await loadRuns();
  }

  return (
    <main className="shell">
      <aside className="sidebar">
        <div className="brand">Workforce Allocation</div>
        <div className="nav-item active">
          <UsersRound size={18} />
          Dashboard
        </div>
        <div className="nav-item">
          <UploadCloud size={18} />
          Upload
        </div>
        <div className="nav-item">
          <Database size={18} />
          Runs
        </div>
      </aside>

      <section className="main">
        <div className="topbar">
          <div>
            <h1 className="title">จัดสรรคนประจำวัน</h1>
            <p className="subtitle">Upload ไฟล์ input แล้วติดตามสถานะการประมวลผล</p>
          </div>
          {user ? (
            <button
              className="button"
              onClick={() => supabase.auth.signOut()}
              type="button"
            >
              Sign out
            </button>
          ) : null}
        </div>

        <div className="grid">
          {!user ? (
            <section className="panel">
              <h2>Login</h2>
              <div className="field">
                <label htmlFor="email">Email</label>
                <input
                  className="file-input"
                  id="email"
                  type="email"
                  value={email}
                  onChange={(event) => setEmail(event.target.value)}
                  placeholder="name@company.com"
                />
              </div>
              <button
                className="button"
                disabled={!email}
                onClick={sendMagicLink}
                type="button"
              >
                Send Magic Link
              </button>
              {message ? <div className="message">{message}</div> : null}
              {error ? <div className="message error">{error}</div> : null}
            </section>
          ) : null}

          <section className="panel">
            <h2>Upload Input</h2>
            {requiredFiles.map((item) => (
              <div className="field" key={item.key}>
                <label htmlFor={item.key}>{item.label}</label>
                <input
                  className="file-input"
                  id={item.key}
                  type="file"
                  accept=".csv,.xlsx,.xls"
                  onChange={(event) =>
                    setFiles((current) => ({
                      ...current,
                      [item.key]: event.target.files?.[0] ?? null,
                    }))
                  }
                />
              </div>
            ))}
            <button
              className="button"
              disabled={!canUpload || isUploading}
              onClick={uploadFiles}
              type="button"
            >
              <UploadCloud size={17} />
              {isUploading ? "Uploading" : "Create Run"}
            </button>
            {message ? <div className="message">{message}</div> : null}
            {error ? <div className="message error">{error}</div> : null}
          </section>

          <section className="panel">
            <h2>Recent Runs</h2>
            <table className="table">
              <thead>
                <tr>
                  <th>Created</th>
                  <th>Status</th>
                  <th>Target Date</th>
                  <th>Solver</th>
                </tr>
              </thead>
              <tbody>
                {runs.map((run) => (
                  <tr key={run.id}>
                    <td>{new Date(run.created_at).toLocaleString("th-TH")}</td>
                    <td>
                      <span className="status">{run.status}</span>
                    </td>
                    <td>{run.target_date ?? "-"}</td>
                    <td>{run.solver_status ?? "-"}</td>
                  </tr>
                ))}
                {runs.length === 0 ? (
                  <tr>
                    <td colSpan={4}>ยังไม่มี run</td>
                  </tr>
                ) : null}
              </tbody>
            </table>
          </section>
        </div>
      </section>
    </main>
  );
}
