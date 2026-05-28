"use client";

import { useEffect, useMemo, useState } from "react";
import {
  BarChart3,
  BriefcaseBusiness,
  CalendarDays,
  ChevronDown,
  ClipboardCheck,
  Database,
  Download,
  FileSpreadsheet,
  Home as HomeIcon,
  LayoutGrid,
  LogOut,
  Settings,
  UploadCloud,
  UsersRound,
} from "lucide-react";
import type { Dispatch, ReactNode, SetStateAction } from "react";
import { supabase } from "@/lib/supabase";

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

type TabId =
  | "dashboard"
  | "timestamp"
  | "run"
  | "results"
  | "timestamp_dept"
  | "master"
  | "skill"
  | "report"
  | "setting";

const navItems = [
  { id: "dashboard", label: "Dashboard", icon: HomeIcon },
  { id: "timestamp", label: "Upload Timestamp", icon: UploadCloud },
  { id: "run", label: "Run Allocation", icon: ClipboardCheck },
  { id: "results", label: "ผลลัพธ์การจัดสรร", icon: BriefcaseBusiness },
  { id: "timestamp_dept", label: "Timestamp With Dept", icon: Database },
  { id: "master", label: "Master Data", icon: FileSpreadsheet },
  { id: "skill", label: "Skill Matrix", icon: LayoutGrid },
  { id: "report", label: "Report & Dashboard", icon: BarChart3 },
  { id: "setting", label: "Setting", icon: Settings },
];

const deptRows = [
  { dept: "ฝ่ายตัดแต่ง", value: 312, percent: 100 },
  { dept: "ฝ่ายผลิต", value: 201, percent: 64 },
  { dept: "ฝ่ายบรรจุ", value: 158, percent: 51 },
  { dept: "ฝ่ายซ่อมบำรุง", value: 95, percent: 30 },
  { dept: "ฝ่ายคลังสินค้า", value: 86, percent: 28 },
];

const allocationRows = [
  ["10012345", "สมชาย ใจดี", "ฝ่ายตัดแต่ง", "พนักงาน", "ตัดแต่งมันหลัง", "4", "06:45"],
  ["10012346", "อารีย์ รักงาน", "ฝ่ายผลิต", "พนักงาน", "ตัดชิ้นส่วน", "3", "06:48"],
  ["10012347", "วิชัย กล้าแข็ง", "ฝ่ายผลิต", "พนักงาน", "ล้างทำความสะอาด", "2", "06:50"],
  ["10012348", "รนพร ขยันดี", "ฝ่ายบรรจุ", "พนักงาน", "บรรจุสุญญากาศ", "3", "06:52"],
  ["10012349", "ประเสริฐ ทองดี", "ฝ่ายซ่อมบำรุง", "ช่างเทคนิค", "ซ่อมบำรุงเครื่องจักร", "4", "06:55"],
];

type MasterFileKey = (typeof masterFileTypes)[number]["key"];
type MasterUploadState = Record<MasterFileKey, File | null>;

const publicWorkspace = "public";

export default function Home() {
  const [activeTab, setActiveTab] = useState<TabId>("dashboard");
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

  const latestRun = runs[0];
  const completedRuns = runs.filter((run) => run.status === "completed").length;
  const pendingRuns = runs.filter((run) => run.status !== "completed").length;
  const totalActivePeople = 852;
  const assignedPeople = latestRun?.status === "completed" ? 806 : completedRuns * 120;
  const noSkillPeople = Math.max(46, pendingRuns * 8);
  const workDate = new Date().toLocaleDateString("th-TH", {
    day: "numeric",
    month: "long",
    year: "numeric",
  });
  const activeNav = navItems.find((item) => item.id === activeTab);

  useEffect(() => {
    void loadDashboard();
  }, []);

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

    for (const item of masterFileTypes) {
      const file = masterUploads[item.key];
      if (!file) continue;

      const fileId = crypto.randomUUID();
      const path = `${publicWorkspace}/masters/${item.key}/${fileId}-${file.name}`;
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
        .is("owner_id", null)
        .eq("file_type", item.key);

      if (deactivateError) {
        setError(deactivateError.message);
        setIsSavingMasters(false);
        return;
      }

      const { error: insertError } = await supabase
        .from("master_data_files")
        .insert({
          owner_id: null,
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
    const scanPath = `${publicWorkspace}/runs/${runId}/timestamp-${timestampFile.name}`;
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
      owner_id: null,
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
    <main className="app-shell">
      <aside className="sidebar">
        <div className="cpf-logo">
          <span>CPF</span>
        </div>

        <nav className="nav-list" aria-label="Main navigation">
          {navItems.map((item) => {
            const Icon = item.icon;
            return (
              <button
                className={`nav-item ${activeTab === item.id ? "active" : ""}`}
                key={item.label}
                onClick={() => setActiveTab(item.id as TabId)}
                type="button"
              >
                <Icon size={19} />
                <span>{item.label}</span>
                {item.label === "Master Data" || item.label === "Report & Dashboard" ? (
                  <ChevronDown className="nav-chevron" size={15} />
                ) : null}
              </button>
            );
          })}
        </nav>

        <div className="logout">
          <LogOut size={19} />
          <span>ออกจากระบบ</span>
        </div>
      </aside>

      <section className="main">
        <header className="topbar">
          <div>
            <h1 className="title">Workforce Allocation System</h1>
            <p className="subtitle">ระบบจัดสรรตำแหน่งงานอัตโนมัติ</p>
          </div>
          <div className="top-actions">
            <div className="date-picker">
              <CalendarDays size={19} />
              <div>
                <span>วันที่ทำงาน</span>
                <strong>{workDate}</strong>
              </div>
              <ChevronDown size={17} />
            </div>
            <div className="admin-chip">
              <div className="avatar">A</div>
              <div>
                <strong>Admin</strong>
                <span>Administrator</span>
              </div>
              <ChevronDown size={17} />
            </div>
          </div>
        </header>

        <section className="dashboard-head">
          <h2>{activeNav?.label ?? "Dashboard"}</h2>
          {(message || error) ? (
            <div className={`toast ${error ? "error" : ""}`}>{error || message}</div>
          ) : null}
        </section>

        {activeTab === "dashboard" ? (
          <>
            <section className="kpi-grid">
              <KpiCard
                icon={<UsersRound size={34} />}
                tone="green"
                label="พนักงานที่มาทำงาน"
                value={totalActivePeople.toLocaleString()}
                unit="คน"
                note="จากทั้งหมด 1,024 คน"
                progress={83}
              />
              <KpiCard
                icon={<ClipboardCheck size={34} />}
                tone="blue"
                label="จัดสรรสำเร็จ"
                value={assignedPeople.toLocaleString()}
                unit="คน"
                note="94.60% ของผู้ที่มาทำงาน"
              />
              <KpiCard
                icon={<UsersRound size={34} />}
                tone="amber"
                label="รอจัดสรร (ไม่มี Skill)"
                value={noSkillPeople.toLocaleString()}
                unit="คน"
                note="5.40% ของผู้ที่มาทำงาน"
              />
              <KpiCard
                icon={<BriefcaseBusiness size={34} />}
                tone="purple"
                label="หน่วยงานทั้งหมด"
                value="12"
                unit="แผนก"
                note="สถานีงานทั้งหมด 68 สถานี"
              />
            </section>

            <DashboardPanels
              activeMasterMap={activeMasterMap}
              assignedPeople={assignedPeople}
              totalActivePeople={totalActivePeople}
            />
          </>
        ) : null}

        {activeTab === "master" ? (
          <MasterDataPage
            activeMasterMap={activeMasterMap}
            canSaveMasters={canSaveMasters}
            isSavingMasters={isSavingMasters}
            saveMasterFiles={saveMasterFiles}
            setMasterUploads={setMasterUploads}
          />
        ) : null}

        {activeTab === "timestamp" ? (
          <TimestampPage
            createDailyRun={createDailyRun}
            hasAllActiveMasters={hasAllActiveMasters}
            isCreatingRun={isCreatingRun}
            setTimestampFile={setTimestampFile}
            timestampFile={timestampFile}
          />
        ) : null}

        {!["dashboard", "master", "timestamp"].includes(activeTab) ? (
          <section className="panel empty-page">
            <h3>{activeNav?.label}</h3>
            <p>แท็บนี้จะเชื่อมข้อมูลจริงในขั้นถัดไป</p>
          </section>
        ) : null}
      </section>
    </main>
  );
}

function DashboardPanels({
  activeMasterMap,
  assignedPeople,
  totalActivePeople,
}: {
  activeMasterMap: Partial<Record<MasterFileKey, MasterFile>>;
  assignedPeople: number;
  totalActivePeople: number;
}) {
  return (
    <>
      <section className="dashboard-grid">
        <section className="panel allocation-status">
          <h3>สถานะการจัดสรรในวันนี้</h3>
          <div className="donut-area">
            <div className="donut">
              <div>
                <strong>{totalActivePeople}</strong>
                <span>พนักงาน</span>
              </div>
            </div>
            <div className="legend">
              <LegendRow color="green" label="จัดสรรสำเร็จ" value="806" percent="94.60%" />
              <LegendRow color="amber" label="ไม่มี Skill" value="46" percent="5.40%" />
              <LegendRow color="gray" label="รอข้อมูล" value="0" percent="0.00%" />
              <LegendRow color="red" label="ขาดงาน" value="172" percent="16.80%" />
            </div>
          </div>
          <p className="panel-note">หมายเหตุ : ขาดงาน คือ พนักงานที่ไม่พบการสแกนเข้างาน</p>
        </section>

        <section className="panel dept-panel">
          <div className="panel-title-row">
            <h3>การจัดสรรตามหน่วยงาน (Top 5)</h3>
            <button className="ghost-button" type="button">ทั้งหมด <ChevronDown size={14} /></button>
          </div>
          <div className="dept-bars">
            {deptRows.map((row) => (
              <div className="dept-row" key={row.dept}>
                <span>{row.dept}</span>
                <div className="bar-track">
                  <div className="bar-fill" style={{ width: `${row.percent}%` }} />
                </div>
                <strong>{row.value}</strong>
              </div>
            ))}
          </div>
          <button className="link-button" type="button">ดูทั้งหมด</button>
        </section>

        <section className="panel files-panel">
          <h3>ไฟล์ล่าสุด</h3>
          <LatestMasterFiles activeMasterMap={activeMasterMap} />
          <button className="link-button" type="button">ดูทั้งหมด</button>
        </section>
      </section>

      <ResultsPanel assignedPeople={assignedPeople} />
    </>
  );
}

function MasterDataPage({
  activeMasterMap,
  canSaveMasters,
  isSavingMasters,
  saveMasterFiles,
  setMasterUploads,
}: {
  activeMasterMap: Partial<Record<MasterFileKey, MasterFile>>;
  canSaveMasters: boolean;
  isSavingMasters: boolean;
  saveMasterFiles: () => Promise<void>;
  setMasterUploads: Dispatch<SetStateAction<MasterUploadState>>;
}) {
  return (
    <section className="workspace-grid">
      <section className="panel master-management">
        <div className="panel-title-row">
          <div>
            <h3>Master Data</h3>
            <p>อัปโหลดไฟล์หลัก 3 ไฟล์ครั้งเดียว แล้วระบบจะใช้ชุดล่าสุดกับ daily run อัตโนมัติ</p>
          </div>
        </div>

        <div className="master-upload-grid">
          {masterFileTypes.map((item) => {
            const activeFile = activeMasterMap[item.key];
            return (
              <label className="master-upload-card" key={item.key}>
                <FileSpreadsheet size={28} />
                <strong>{item.label}</strong>
                <span>{activeFile?.original_filename ?? "ยังไม่มีไฟล์ active"}</span>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(event) =>
                    setMasterUploads((current) => ({
                      ...current,
                      [item.key]: event.target.files?.[0] ?? null,
                    }))
                  }
                />
              </label>
            );
          })}
        </div>

        <button
          className="primary-button save-master-button"
          disabled={!canSaveMasters || isSavingMasters}
          onClick={saveMasterFiles}
          type="button"
        >
          <UploadCloud size={18} />
          {isSavingMasters ? "Saving" : "Save Master Files"}
        </button>
      </section>

      <section className="panel files-panel">
        <h3>Active Master Files</h3>
        <LatestMasterFiles activeMasterMap={activeMasterMap} />
      </section>
    </section>
  );
}

function TimestampPage({
  createDailyRun,
  hasAllActiveMasters,
  isCreatingRun,
  setTimestampFile,
  timestampFile,
}: {
  createDailyRun: () => Promise<void>;
  hasAllActiveMasters: boolean;
  isCreatingRun: boolean;
  setTimestampFile: (file: File | null) => void;
  timestampFile: File | null;
}) {
  return (
    <section className="workspace-grid">
      <section className="panel master-management">
        <h3>Upload Timestamp</h3>
        <p>อัปโหลดเฉพาะไฟล์ timestamp รายวัน ระบบจะใช้ master data ชุด active ล่าสุด</p>
        <label className="master-upload-card single">
          <UploadCloud size={30} />
          <strong>{timestampFile?.name ?? "Timestamp / Time Record"}</strong>
          <span>รองรับ .csv, .xlsx, .xls</span>
          <input
            type="file"
            accept=".csv,.xlsx,.xls"
            onChange={(event) => setTimestampFile(event.target.files?.[0] ?? null)}
          />
        </label>
        <button
          className="primary-button save-master-button"
          disabled={!timestampFile || !hasAllActiveMasters || isCreatingRun}
          onClick={createDailyRun}
          type="button"
        >
          <ClipboardCheck size={18} />
          {isCreatingRun ? "Creating" : "Create Daily Run"}
        </button>
        {!hasAllActiveMasters ? (
          <p className="inline-warning">ต้องมี master files ครบ 3 ไฟล์ก่อนสร้าง daily run</p>
        ) : null}
      </section>
    </section>
  );
}

function LatestMasterFiles({
  activeMasterMap,
}: {
  activeMasterMap: Partial<Record<MasterFileKey, MasterFile>>;
}) {
  return (
    <div className="file-stack">
      {masterFileTypes.map((item) => {
        const activeFile = activeMasterMap[item.key];
        return (
          <div className="file-card" key={item.key}>
            <FileSpreadsheet size={24} />
            <div>
              <strong>{activeFile?.original_filename ?? item.label}</strong>
              <span>
                {activeFile
                  ? new Date(activeFile.created_at).toLocaleString("th-TH")
                  : "ยังไม่มีไฟล์"}
              </span>
            </div>
            <Download size={18} />
          </div>
        );
      })}
    </div>
  );
}

function ResultsPanel({ assignedPeople }: { assignedPeople: number }) {
  return (
    <section className="panel results-panel">
      <div className="panel-title-row">
        <h3>ผลลัพธ์การจัดสรรล่าสุด</h3>
        <div className="table-actions">
          <button className="ghost-button" type="button">ดูทั้งหมด</button>
          <button className="primary-button small" type="button">
            Export <ChevronDown size={15} />
          </button>
        </div>
      </div>
      <table className="table">
        <thead>
          <tr>
            <th>No.</th>
            <th>รหัสพนักงาน</th>
            <th>ชื่อ-สกุล</th>
            <th>หน่วยงาน</th>
            <th>ตำแหน่ง</th>
            <th>สถานีงานที่จัดสรร</th>
            <th>ระดับ Skill</th>
            <th>เวลาเข้า</th>
            <th>สถานะ</th>
          </tr>
        </thead>
        <tbody>
          {allocationRows.map((row, index) => (
            <tr key={row[0]}>
              <td>{index + 1}</td>
              {row.map((cell) => (
                <td key={`${row[0]}-${cell}`}>{cell}</td>
              ))}
              <td><span className="status-pill">จัดสรรสำเร็จ</span></td>
            </tr>
          ))}
        </tbody>
      </table>
      <div className="pagination">
        <span>1-5 จาก {assignedPeople || 806} รายการ</span>
        <button type="button">1</button>
        <button type="button">2</button>
        <button type="button">3</button>
        <button type="button">...</button>
      </div>
    </section>
  );
}

function KpiCard({
  icon,
  tone,
  label,
  value,
  unit,
  note,
  progress,
}: {
  icon: ReactNode;
  tone: "green" | "blue" | "amber" | "purple";
  label: string;
  value: string;
  unit: string;
  note: string;
  progress?: number;
}) {
  return (
    <article className="kpi-card">
      <div className={`kpi-icon ${tone}`}>{icon}</div>
      <div className="kpi-body">
        <span>{label}</span>
        <div>
          <strong>{value}</strong>
          <b>{unit}</b>
        </div>
        <p>{note}</p>
        {progress ? (
          <div className="progress-line">
            <i style={{ width: `${progress}%` }} />
            <em>{progress}.20%</em>
          </div>
        ) : null}
      </div>
    </article>
  );
}

function LegendRow({
  color,
  label,
  value,
  percent,
}: {
  color: "green" | "amber" | "gray" | "red";
  label: string;
  value: string;
  percent: string;
}) {
  return (
    <div className="legend-row">
      <span className={`dot ${color}`} />
      <p>{label}</p>
      <strong>{value}</strong>
      <em>{percent}</em>
    </div>
  );
}
