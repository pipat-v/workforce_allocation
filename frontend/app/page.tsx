"use client";

import { useEffect, useMemo, useState } from "react";
import { Database, FileSpreadsheet, UploadCloud, UsersRound } from "lucide-react";
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

type MasterFile = {
  id: string;
  file_type: MasterFileKey;
  file_path: string;
  original_filename: string | null;
  created_at: string;
};

const masterFileTypes = [
  { key: "employee_master", label: "รายชื่อพนักงาน" },
  { key: "manpower_plan", label: "Manpower Plan" },
  { key: "skill_matrix", label: "Skill Matrix" },
] as const;

type MasterFileKey = (typeof masterFileTypes)[number]["key"];
type MasterUploadState = Record<MasterFileKey, File | null>;

export default function Home() {
  const [user, setUser] = useState<User | null>(null);
  const [email, setEmail] = useState("");
  const [masterUploads, setMasterUploads] = useState<MasterUploadState>({
    employee_master: null,
    manpower_plan: null,
    skill_matrix: null,
  });
  const [activeMasters, setActiveMasters] = useState<MasterFile[]>([]);
  const [timestampFile, setTimestampFile] = useState<File | null>(null);
  const [runs, setRuns] = useState<AllocationRun[]>([]);
  const [message, setMessage] = useState("");
  const [error, setError] = useState("");
  const [isSavingMasters, setIsSavingMasters] = useState(false);
  const [isCreatingRun, setIsCreatingRun] = useState(false);

  const activeMasterMap = useMemo(
    () =>
      activeMasters.reduce(
        (current, item) => ({ ...current, [item.file_type]: item }),
        {} as Partial<Record<MasterFileKey, MasterFile>>,
      ),
    [activeMasters],
  );

  const hasAllActiveMasters = useMemo(
    () => masterFileTypes.every((item) => activeMasterMap[item.key]),
    [activeMasterMap],
  );

  const canSaveMasters = useMemo(
    () => masterFileTypes.some((item) => masterUploads[item.key]),
    [masterUploads],
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

    return () => {
      subscription.subscription.unsubscribe();
    };
  }, []);

  useEffect(() => {
    if (user) {
      void loadDashboard();
      return;
    }

    setActiveMasters([]);
    setRuns([]);
  }, [user]);

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

  async function loadDashboard() {
    await Promise.all([loadRuns(), loadActiveMasters()]);
  }

  async function loadActiveMasters() {
    const { data, error: loadError } = await supabase
      .from("master_data_files")
      .select("id,file_type,file_path,original_filename,created_at")
      .eq("is_active", true)
      .order("created_at", { ascending: false });

    if (loadError) {
      setError(loadError.message);
      return;
    }

    const latestByType = new Map<MasterFileKey, MasterFile>();
    for (const item of (data ?? []) as MasterFile[]) {
      if (!latestByType.has(item.file_type)) {
        latestByType.set(item.file_type, item);
      }
    }

    setActiveMasters(Array.from(latestByType.values()));
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

  async function saveMasterFiles() {
    setError("");
    setMessage("");
    setIsSavingMasters(true);

    if (!user) {
      setError("กรุณา login ก่อน upload ไฟล์");
      setIsSavingMasters(false);
      return;
    }

    for (const item of masterFileTypes) {
      const file = masterUploads[item.key];
      if (!file) continue;

      const fileId = crypto.randomUUID();
      const path = `${user.id}/masters/${item.key}/${fileId}-${file.name}`;
      const { error: uploadError } = await supabase.storage
        .from("workforce-inputs")
        .upload(path, file, { upsert: true });

      if (uploadError) {
        setError(uploadError.message);
        setIsSavingMasters(false);
        return;
      }

      const { error: deactivateError } = await supabase
        .from("master_data_files")
        .update({ is_active: false })
        .eq("owner_id", user.id)
        .eq("file_type", item.key);

      if (deactivateError) {
        setError(deactivateError.message);
        setIsSavingMasters(false);
        return;
      }

      const { error: insertError } = await supabase
        .from("master_data_files")
        .insert({
          owner_id: user.id,
          file_type: item.key,
          file_path: path,
          original_filename: file.name,
          is_active: true,
        });

      if (insertError) {
        setError(insertError.message);
        setIsSavingMasters(false);
        return;
      }
    }

    setMasterUploads({
      employee_master: null,
      manpower_plan: null,
      skill_matrix: null,
    });
    setMessage("บันทึก master files แล้ว");
    setIsSavingMasters(false);
    await loadActiveMasters();
  }

  async function createDailyRun() {
    setError("");
    setMessage("");
    setIsCreatingRun(true);

    if (!user) {
      setError("กรุณา login ก่อนสร้าง run");
      setIsCreatingRun(false);
      return;
    }

    if (!timestampFile) {
      setError("กรุณาเลือกไฟล์ timestamp");
      setIsCreatingRun(false);
      return;
    }

    if (!hasAllActiveMasters) {
      setError("กรุณา upload master files ให้ครบก่อนสร้าง daily run");
      setIsCreatingRun(false);
      return;
    }

    const runId = crypto.randomUUID();
    const scanPath = `${user.id}/runs/${runId}/timestamp-${timestampFile.name}`;
    const { error: uploadError } = await supabase.storage
      .from("workforce-inputs")
      .upload(scanPath, timestampFile, { upsert: true });

    if (uploadError) {
      setError(uploadError.message);
      setIsCreatingRun(false);
      return;
    }

    const { error: insertError } = await supabase.from("allocation_runs").insert({
      id: runId,
      owner_id: user.id,
      status: "uploaded",
      scan_file_path: scanPath,
      master_file_path: activeMasterMap.employee_master?.file_path,
      manpower_file_path: activeMasterMap.manpower_plan?.file_path,
      skill_file_path: activeMasterMap.skill_matrix?.file_path,
    });

    if (insertError) {
      setError(insertError.message);
      setIsCreatingRun(false);
      return;
    }

    setTimestampFile(null);
    setMessage("สร้าง daily run แล้ว รอ worker ประมวลผล");
    setIsCreatingRun(false);
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
            <h2>Master Files</h2>
            <div className="master-list">
              {masterFileTypes.map((item) => {
                const activeFile = activeMasterMap[item.key];
                return (
                  <div className="master-row" key={item.key}>
                    <FileSpreadsheet size={18} />
                    <div>
                      <strong>{item.label}</strong>
                      <span>{activeFile?.original_filename ?? "ยังไม่มีไฟล์"}</span>
                    </div>
                  </div>
                );
              })}
            </div>
            {masterFileTypes.map((item) => (
              <div className="field" key={item.key}>
                <label htmlFor={item.key}>{item.label}</label>
                <input
                  className="file-input"
                  id={item.key}
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(event) =>
                    setMasterUploads((current) => ({
                      ...current,
                      [item.key]: event.target.files?.[0] ?? null,
                    }))
                  }
                />
              </div>
            ))}
            <button
              className="button"
              disabled={!canSaveMasters || isSavingMasters}
              onClick={saveMasterFiles}
              type="button"
            >
              <UploadCloud size={17} />
              {isSavingMasters ? "Saving" : "Save Master Files"}
            </button>
            {message ? <div className="message">{message}</div> : null}
            {error ? <div className="message error">{error}</div> : null}
          </section>

          <section className="panel">
            <h2>Daily Timestamp</h2>
            <div className="field">
              <label htmlFor="timestamp">Timestamp / Time Record</label>
              <input
                className="file-input"
                id="timestamp"
                type="file"
                accept=".csv,.xlsx,.xls"
                onChange={(event) =>
                  setTimestampFile(event.target.files?.[0] ?? null)
                }
              />
            </div>
            <button
              className="button"
              disabled={!timestampFile || !hasAllActiveMasters || isCreatingRun}
              onClick={createDailyRun}
              type="button"
            >
              <UploadCloud size={17} />
              {isCreatingRun ? "Creating" : "Create Daily Run"}
            </button>
            {!hasAllActiveMasters ? (
              <div className="message error">
                ต้องมี master files ครบ 3 ไฟล์ก่อนสร้าง daily run
              </div>
            ) : null}
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
