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
import * as XLSX from "xlsx";
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

type AttendanceRecord = {
  empId: string;
  name: string;
  dept: string;
  position: string;
  shift: string;
  shiftStart: string;
  scanIn: string;
  status: "Present" | "Late" | "Absent";
  minutesLate: number;
};

type ReportData = {
  targetDate: string;
  totalEmployees: number;
  present: number;
  late: number;
  absent: number;
  deptRows: Array<{ dept: string; present: number; late: number; absent: number; total: number }>;
  lateRows: AttendanceRecord[];
  records: AttendanceRecord[];
  timestampRows: AttendanceRecord[];
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
  const [reportData, setReportData] = useState<ReportData | null>(null);
  const [loadedReportKey, setLoadedReportKey] = useState("");
  const [isLoadingReport, setIsLoadingReport] = useState(false);
  const [isSavingMasters, setIsSavingMasters] = useState(false);
  const [isCreatingRun, setIsCreatingRun] = useState(false);
  const [resultsPage, setResultsPage] = useState(1);
  const [timestampDeptPage, setTimestampDeptPage] = useState(1);

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
  const reportSourceKey =
    activeMasterMap.employee_master?.file_path && latestRun?.scan_file_path
      ? [
          activeMasterMap.employee_master.file_path,
          activeMasterMap.manpower_plan?.file_path,
          latestRun.scan_file_path,
        ].filter(Boolean).join("|")
      : "";
  const totalEmployees = reportData?.totalEmployees ?? 0;
  const presentPeople = reportData?.present ?? 0;
  const latePeople = reportData?.late ?? 0;
  const absentPeople = reportData?.absent ?? 0;
  const totalActivePeople = presentPeople + latePeople;
  const presentRate = totalEmployees ? Math.round((totalActivePeople / totalEmployees) * 100) : 0;
  const workDate = new Date().toLocaleDateString("th-TH", {
    day: "numeric",
    month: "long",
    year: "numeric",
  });
  const activeNav = navItems.find((item) => item.id === activeTab);

  useEffect(() => {
    void loadDashboard();
  }, []);

  useEffect(() => {
    if (!reportSourceKey || loadedReportKey === reportSourceKey || isLoadingReport) {
      return;
    }

    void loadReportDashboard();
  }, [reportSourceKey, loadedReportKey, isLoadingReport]);

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
      const path = `${publicWorkspace}/masters/${item.key}/${fileId}${getSafeFileExtension(file.name)}`;
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
    const scanPath = `${publicWorkspace}/runs/${runId}/timestamp${getSafeFileExtension(timestampFile.name)}`;
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

  async function loadReportDashboard() {
    setError("");
    setMessage("");
    setIsLoadingReport(true);

    try {
      const employeeMaster = activeMasterMap.employee_master;
      const latestRun = runs.find((run) => run.scan_file_path);

      if (!employeeMaster || !latestRun?.scan_file_path) {
        setError("ต้องมีไฟล์รายชื่อพนักงานและ timestamp ล่าสุดก่อนสร้าง Report & Dashboard");
        setIsLoadingReport(false);
        return;
      }

      const [employeeRows, scanRows, manpowerRows] = await Promise.all([
        downloadSheetRows(employeeMaster.file_path),
        downloadSheetRows(latestRun.scan_file_path),
        activeMasterMap.manpower_plan
          ? downloadSheetRows(activeMasterMap.manpower_plan.file_path)
          : Promise.resolve([]),
      ]);

      setReportData(buildReportData(employeeRows, scanRows, manpowerRows));
      setLoadedReportKey([
        employeeMaster.file_path,
        activeMasterMap.manpower_plan?.file_path,
        latestRun.scan_file_path,
      ].filter(Boolean).join("|"));
    } catch (reportError) {
      setError(reportError instanceof Error ? reportError.message : "โหลด report ไม่สำเร็จ");
    } finally {
      setIsLoadingReport(false);
    }
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="cpf-logo">
          <img alt="CPF" src="/cpf-logo.png" />
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
                note={`จากทั้งหมด ${totalEmployees.toLocaleString()} คน`}
                progress={presentRate}
              />
              <KpiCard
                icon={<ClipboardCheck size={34} />}
                tone="blue"
                label="Present"
                value={presentPeople.toLocaleString()}
                unit="คน"
                note="พนักงานที่เข้างานตรงเวลา"
              />
              <KpiCard
                icon={<UsersRound size={34} />}
                tone="amber"
                label="Late"
                value={latePeople.toLocaleString()}
                unit="คน"
                note="พนักงานที่มาสาย"
              />
              <KpiCard
                icon={<BriefcaseBusiness size={34} />}
                tone="purple"
                label="Absent"
                value={absentPeople.toLocaleString()}
                unit="คน"
                note="ไม่พบการสแกนเข้างาน"
              />
            </section>

            <DashboardPanels
              activeMasterMap={activeMasterMap}
              assignedPeople={presentPeople}
              reportData={reportData}
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

        {activeTab === "run" ? (
          <RunAllocationPage
            createDailyRun={createDailyRun}
            hasAllActiveMasters={hasAllActiveMasters}
            isCreatingRun={isCreatingRun}
            latestRun={latestRun}
            setTimestampFile={setTimestampFile}
            timestampFile={timestampFile}
          />
        ) : null}

        {activeTab === "results" ? (
          <ResultsPanel
            page={resultsPage}
            reportData={reportData}
            setPage={setResultsPage}
            standalone
          />
        ) : null}

        {activeTab === "timestamp_dept" ? (
          <TimestampWithDeptPage
            page={timestampDeptPage}
            reportData={reportData}
            setPage={setTimestampDeptPage}
          />
        ) : null}

        {activeTab === "report" ? (
          <ReportDashboard
            isLoadingReport={isLoadingReport}
            loadReportDashboard={loadReportDashboard}
            reportData={reportData}
          />
        ) : null}

        {!["dashboard", "master", "timestamp", "run", "results", "timestamp_dept", "report"].includes(activeTab) ? (
          <section className="panel empty-page">
            <h3>{activeNav?.label}</h3>
            <p>แท็บนี้จะเชื่อมข้อมูลจริงในขั้นถัดไป</p>
          </section>
        ) : null}
      </section>
    </main>
  );
}

async function downloadSheetRows(path: string): Promise<Record<string, unknown>[]> {
  const { data, error } = await supabase.storage.from("workforce-inputs").download(path);
  if (error) {
    throw new Error(error.message);
  }

  const buffer = await data.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(firstSheet, {
    defval: "",
    raw: false,
  });

  if (rows.some((row) => "Timestamp" in row && "Employee ID" in row)) {
    return rows;
  }

  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, {
    header: 1,
    defval: "",
    raw: false,
  });
  const headerIndex = rawRows.findIndex(
    (row) =>
      Array.isArray(row) &&
      row.includes("Timestamp") &&
      (row.includes("Employee ID") || row.includes("Employee Name")),
  );

  if (headerIndex < 0) {
    return rows;
  }

  const headers = rawRows[headerIndex].map((header) => String(header));
  return rawRows.slice(headerIndex + 1).map((row) => {
    const record: Record<string, unknown> = {};
    headers.forEach((header, index) => {
      record[header] = Array.isArray(row) ? row[index] ?? "" : "";
    });
    return record;
  });
}

function buildReportData(
  employeeRows: Record<string, unknown>[],
  scanRows: Record<string, unknown>[],
  manpowerRows: Record<string, unknown>[],
): ReportData {
  const employees = employeeRows.map((row) => {
    const empId = cleanEmpId(row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"]);
    const firstName = String(row["First Name (Local)"] ?? "").trim();
    const lastName = String(row["Last Name (Local)"] ?? "").trim();
    const fallbackName = String(row["Employee Name"] ?? row["Name"] ?? "").trim();
    return {
      empId,
      name: `${firstName} ${lastName}`.trim() || fallbackName || empId,
      dept: String(row["หน่วยงาน"] ?? row["dept"] ?? row["Name (Section)"] ?? "ไม่ระบุ").trim() || "ไม่ระบุ",
      position: String(row["Title (Position)"] ?? row["position"] ?? row["ตำแหน่ง"] ?? "พนักงาน").trim() || "พนักงาน",
    };
  }).filter((row) => row.empId);

  const employeeMap = new Map(employees.map((employee) => [employee.empId, employee]));
  const deptShiftStart = buildDeptShiftStart(manpowerRows);
  const scanByEmp = new Map<string, { name: string; times: Date[] }>();

  for (const row of scanRows) {
    const empId = cleanEmpId(row["Employee ID"] ?? row["Emp ID"] ?? row["รหัสพนักงาน"]);
    const timestamp = parseTimestamp(row["Timestamp"]);
    if (!empId || !timestamp) continue;

    const current = scanByEmp.get(empId) ?? {
      name: String(row["Employee Name"] ?? row["name"] ?? "").trim(),
      times: [],
    };
    current.times.push(timestamp);
    scanByEmp.set(empId, current);
  }

  const targetDate =
    Array.from(scanByEmp.values())
      .flatMap((entry) => entry.times)
      .sort((a, b) => b.getTime() - a.getTime())[0]
      ?.toLocaleDateString("th-TH") ?? "-";

  const baseEmployees = employees.length
    ? employees
    : Array.from(scanByEmp.entries()).map(([empId, entry]) => ({
        empId,
        name: entry.name || empId,
        dept: "ไม่ระบุ",
        position: "พนักงาน",
      }));

  const records: AttendanceRecord[] = baseEmployees.map((employee) => {
    const scans = scanByEmp.get(employee.empId)?.times ?? [];
    const scanIn = scans.sort((a, b) => a.getTime() - b.getTime())[0];
    const shiftStart = deptShiftStart.get(employee.dept) ?? "07:00";
    const minutesLate = scanIn ? Math.max(0, minutesBetween(shiftStart, scanIn)) : 0;
    const status = !scanIn ? "Absent" : minutesLate > 5 ? "Late" : "Present";

    return {
      empId: employee.empId,
      name: employee.name,
      dept: employee.dept,
      position: employee.position,
      shift: "กะ 1",
      shiftStart,
      scanIn: scanIn ? toTimeText(scanIn) : "-",
      status,
      minutesLate,
    };
  });

  const deptMap = new Map<string, { dept: string; present: number; late: number; absent: number; total: number }>();
  for (const record of records) {
    const current = deptMap.get(record.dept) ?? {
      dept: record.dept,
      present: 0,
      late: 0,
      absent: 0,
      total: 0,
    };
    current.total += 1;
    if (record.status === "Present") current.present += 1;
    if (record.status === "Late") current.late += 1;
    if (record.status === "Absent") current.absent += 1;
    deptMap.set(record.dept, current);
  }

  return {
    targetDate,
    totalEmployees: records.length,
    present: records.filter((record) => record.status === "Present").length,
    late: records.filter((record) => record.status === "Late").length,
    absent: records.filter((record) => record.status === "Absent").length,
    deptRows: Array.from(deptMap.values())
      .sort((a, b) => b.total - a.total)
      .slice(0, 12),
    lateRows: records
      .filter((record) => record.status === "Late")
      .sort((a, b) => b.minutesLate - a.minutesLate)
      .slice(0, 80),
    records,
    timestampRows: records.filter((record) => record.scanIn !== "-"),
  };
}

function buildDeptShiftStart(rows: Record<string, unknown>[]) {
  const map = new Map<string, string>();
  for (const row of rows) {
    const dept = String(row["หน่วยงาน"] ?? row["dept"] ?? "").trim();
    const shiftStart = normalizeTimeText(row["เวลาเข้า"] ?? row["shift_start"]);
    if (dept && shiftStart && !map.has(dept)) {
      map.set(dept, shiftStart);
    }
  }
  return map;
}

function cleanEmpId(value: unknown) {
  return String(value ?? "")
    .replace(/\u00a0/g, "")
    .replace(/\s+/g, "")
    .replace(/\.0$/, "")
    .replace(/[^0-9]/g, "");
}

function parseTimestamp(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  const text = String(value ?? "").trim();
  const thaiStyle = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (thaiStyle) {
    const [, day, month, year, hour, minute, second = "0"] = thaiStyle;
    return new Date(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour),
      Number(minute),
      Number(second),
    );
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function normalizeTimeText(value: unknown) {
  if (value instanceof Date) {
    return toTimeText(value);
  }

  const text = String(value ?? "").trim();
  const match = text.match(/(\d{1,2}):(\d{2})/);
  if (!match) return "";
  return `${match[1].padStart(2, "0")}:${match[2]}`;
}

function toTimeText(value: Date) {
  return value.toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
}

function minutesBetween(shiftStart: string, scanIn: Date) {
  const [hour, minute] = shiftStart.split(":").map(Number);
  const shift = new Date(scanIn);
  shift.setHours(hour || 0, minute || 0, 0, 0);
  return Math.round((scanIn.getTime() - shift.getTime()) / 60000);
}

function getSafeFileExtension(filename: string) {
  const match = filename.toLowerCase().match(/\.(csv|xlsx|xls)$/);
  return match ? `.${match[1]}` : "";
}

function DashboardPanels({
  activeMasterMap,
  assignedPeople,
  reportData,
  totalActivePeople,
}: {
  activeMasterMap: Partial<Record<MasterFileKey, MasterFile>>;
  assignedPeople: number;
  reportData: ReportData | null;
  totalActivePeople: number;
}) {
  const total = reportData?.totalEmployees ?? 0;
  const present = reportData?.present ?? 0;
  const late = reportData?.late ?? 0;
  const absent = reportData?.absent ?? 0;
  const presentPercent = total ? ((present / total) * 100).toFixed(1) : "0.0";
  const latePercent = total ? ((late / total) * 100).toFixed(1) : "0.0";
  const absentPercent = total ? ((absent / total) * 100).toFixed(1) : "0.0";
  const topDeptRows = reportData?.deptRows?.slice(0, 5) ?? deptRows;
  const maxDeptTotal = Math.max(...topDeptRows.map((row) => "total" in row ? row.total : row.value), 1);

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
              <LegendRow color="green" label="Present" value={String(present)} percent={`${presentPercent}%`} />
              <LegendRow color="amber" label="Late" value={String(late)} percent={`${latePercent}%`} />
              <LegendRow color="gray" label="รอข้อมูล" value="0" percent="0.0%" />
              <LegendRow color="red" label="Absent" value={String(absent)} percent={`${absentPercent}%`} />
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
            {topDeptRows.map((row) => {
              const value = "total" in row ? row.total : row.value;
              return (
                <div className="dept-row" key={row.dept}>
                  <span>{row.dept}</span>
                  <div className="bar-track">
                    <div className="bar-fill" style={{ width: `${(value / maxDeptTotal) * 100}%` }} />
                  </div>
                  <strong>{value}</strong>
                </div>
              );
            })}
          </div>
          <button className="link-button" type="button">ดูทั้งหมด</button>
        </section>

        <section className="panel files-panel">
          <h3>ไฟล์ล่าสุด</h3>
          <LatestMasterFiles activeMasterMap={activeMasterMap} />
          <button className="link-button" type="button">ดูทั้งหมด</button>
        </section>
      </section>

      <ResultsPanel reportData={reportData} />
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
    <section className="workspace-grid master-page">
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

function RunAllocationPage({
  createDailyRun,
  hasAllActiveMasters,
  isCreatingRun,
  latestRun,
  setTimestampFile,
  timestampFile,
}: {
  createDailyRun: () => Promise<void>;
  hasAllActiveMasters: boolean;
  isCreatingRun: boolean;
  latestRun?: AllocationRun;
  setTimestampFile: (file: File | null) => void;
  timestampFile: File | null;
}) {
  return (
    <section className="workspace-grid">
      <section className="panel run-panel">
        <h3>Run Allocation</h3>
        <p>เลือก timestamp ล่าสุดเพื่อสร้าง run จาก master data ชุด active</p>
        <label className="master-upload-card single">
          <ClipboardCheck size={30} />
          <strong>{timestampFile?.name ?? "เลือก Timestamp สำหรับ run ใหม่"}</strong>
          <span>ระบบจะผูก master files ล่าสุดเข้ากับ run นี้</span>
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
          {isCreatingRun ? "Creating" : "Create Allocation Run"}
        </button>
        {!hasAllActiveMasters ? (
          <p className="inline-warning">ต้องมี master files ครบ 3 ไฟล์ก่อน</p>
        ) : null}
      </section>

      <section className="panel files-panel">
        <h3>Latest Run</h3>
        {latestRun ? (
          <div className="run-summary">
            <strong>{latestRun.status}</strong>
            <span>{new Date(latestRun.created_at).toLocaleString("th-TH")}</span>
            <span>{latestRun.scan_file_path ?? "-"}</span>
          </div>
        ) : (
          <p className="empty-copy">ยังไม่มี run</p>
        )}
      </section>
    </section>
  );
}

const pageSize = 10;

function TimestampWithDeptPage({
  page,
  reportData,
  setPage,
}: {
  page: number;
  reportData: ReportData | null;
  setPage: (page: number) => void;
}) {
  const allRows = reportData?.timestampRows ?? [];
  const totalPages = Math.max(1, Math.ceil(allRows.length / pageSize));
  const safePage = Math.min(page, totalPages);
  const rows = allRows.slice((safePage - 1) * pageSize, safePage * pageSize);

  return (
    <section className="panel results-panel">
      <div className="panel-title-row">
        <h3>Timestamp With Dept</h3>
        <span className="table-count">{allRows.length} rows</span>
      </div>
      <div className="table-scroll">
        <table className="table data-table">
          <thead>
            <tr>
              <th>Employee ID</th>
              <th>Name</th>
              <th>Dept</th>
              <th>Position</th>
              <th>Shift</th>
              <th>Shift Start</th>
              <th>Scan In</th>
              <th>Status</th>
              <th>Minutes Late</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={`${row.empId}-${row.scanIn}`}>
                <td>{row.empId}</td>
                <td>{row.name}</td>
                <td>{row.dept}</td>
                <td>{row.position}</td>
                <td>{row.shift}</td>
                <td>{row.shiftStart}</td>
                <td>{row.scanIn}</td>
                <td><span className={`status-pill ${row.status.toLowerCase()}`}>{row.status}</span></td>
                <td>{row.minutesLate}</td>
              </tr>
            ))}
            {rows.length === 0 ? (
              <tr>
                <td colSpan={9}>ยังไม่มี timestamp ที่ merge หน่วยงานแล้ว</td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>
      <TablePagination
        page={safePage}
        pageSize={pageSize}
        setPage={setPage}
        totalRows={allRows.length}
      />
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

function ResultsPanel({
  page = 1,
  reportData,
  setPage,
  standalone = false,
}: {
  page?: number;
  reportData: ReportData | null;
  setPage?: (page: number) => void;
  standalone?: boolean;
}) {
  const allRows = reportData?.records.filter((record) => record.status !== "Absent") ?? [];
  const totalPages = Math.max(1, Math.ceil(allRows.length / pageSize));
  const safePage = Math.min(page, totalPages);
  const rows = allRows.slice((safePage - 1) * pageSize, safePage * pageSize);

  return (
    <section className={`panel results-panel ${standalone ? "standalone" : ""}`}>
      <div className="panel-title-row">
        <h3>ผลลัพธ์การจัดสรรล่าสุด</h3>
        <div className="table-actions">
          <button className="ghost-button" type="button">ดูทั้งหมด</button>
          <button className="primary-button small" type="button">
            Export <ChevronDown size={15} />
          </button>
        </div>
      </div>
      <div className="table-scroll">
        <table className="table data-table">
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
            {rows.map((row, index) => (
              <tr key={`${row.empId}-${row.scanIn}`}>
                <td>{(safePage - 1) * pageSize + index + 1}</td>
                <td>{row.empId}</td>
                <td>{row.name}</td>
                <td>{row.dept}</td>
                <td>{row.position}</td>
                <td>{row.dept}</td>
                <td>-</td>
                <td>{row.scanIn}</td>
                <td>
                  <span className={`status-pill ${row.status.toLowerCase()}`}>
                    {row.status}
                  </span>
                </td>
              </tr>
            ))}
            {rows.length === 0 ? (
              <tr>
                <td colSpan={9}>ยังไม่มีข้อมูลจากไฟล์ที่อัปโหลด</td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>
      <TablePagination
        page={safePage}
        pageSize={pageSize}
        setPage={setPage}
        totalRows={allRows.length}
      />
    </section>
  );
}

function TablePagination({
  page,
  pageSize,
  setPage,
  totalRows,
}: {
  page: number;
  pageSize: number;
  setPage?: (page: number) => void;
  totalRows: number;
}) {
  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));
  const start = totalRows === 0 ? 0 : (page - 1) * pageSize + 1;
  const end = Math.min(page * pageSize, totalRows);
  const pages = Array.from({ length: Math.min(totalPages, 5) }, (_, index) => index + 1);

  return (
    <div className="pagination">
      <span>{start}-{end} จาก {totalRows} รายการ</span>
      <button disabled={!setPage || page <= 1} onClick={() => setPage?.(page - 1)} type="button">‹</button>
      {pages.map((item) => (
        <button
          className={item === page ? "active" : ""}
          disabled={!setPage}
          key={item}
          onClick={() => setPage?.(item)}
          type="button"
        >
          {item}
        </button>
      ))}
      {totalPages > 5 ? <span>...</span> : null}
      <button disabled={!setPage || page >= totalPages} onClick={() => setPage?.(page + 1)} type="button">›</button>
    </div>
  );
}

function ReportDashboard({
  isLoadingReport,
  loadReportDashboard,
  reportData,
}: {
  isLoadingReport: boolean;
  loadReportDashboard: () => Promise<void>;
  reportData: ReportData | null;
}) {
  const data = reportData ?? {
    targetDate: "-",
    totalEmployees: 0,
    present: 0,
    late: 0,
    absent: 0,
    deptRows: [],
    lateRows: [],
    records: [],
    timestampRows: [],
  };
  const lateRate = data.totalEmployees
    ? ((data.late / data.totalEmployees) * 100).toFixed(1)
    : "0.0";
  const maxDeptTotal = Math.max(...data.deptRows.map((row) => row.total), 1);
  const presentPercent = data.totalEmployees ? (data.present / data.totalEmployees) * 100 : 0;
  const latePercent = data.totalEmployees ? (data.late / data.totalEmployees) * 100 : 0;
  const absentPercent = data.totalEmployees ? (data.absent / data.totalEmployees) * 100 : 0;

  return (
    <section className="report-page">
      <div className="report-toolbar">
        <div className="filter-card">
          <label>date</label>
          <strong>{data.targetDate}</strong>
        </div>
        <div className="filter-card wide">
          <label>dept</label>
          <strong>{data.deptRows[0]?.dept ?? "Select all"}</strong>
        </div>
        <button
          className="primary-button report-refresh"
          disabled={isLoadingReport}
          onClick={loadReportDashboard}
          type="button"
        >
          <BarChart3 size={17} />
          {isLoadingReport ? "Loading" : "Load Uploaded Data"}
        </button>
      </div>

      <section className="report-kpis">
        <ReportMetric value={data.totalEmployees} label="จำนวนพนักงานทั้งหมด" />
        <ReportMetric value={data.present} label="Present" tone="green" />
        <ReportMetric value={data.late} label="Late" tone="amber" />
        <ReportMetric value={data.absent} label="Absent" tone="red" />
        <ReportMetric value={`${lateRate} %`} label="Late Rate %" />
      </section>

      <section className="report-grid">
        <div className="panel report-card">
          <h3>การเข้างานรายแผนก</h3>
          <div className="stacked-bars">
            {data.deptRows.map((row) => (
              <div className="stacked-row" key={row.dept}>
                <span>{row.dept}</span>
                <div className="stacked-track">
                  <i
                    className="present"
                    style={{ width: `${(row.present / maxDeptTotal) * 100}%` }}
                  />
                  <i
                    className="late"
                    style={{ width: `${(row.late / maxDeptTotal) * 100}%` }}
                  />
                  <i
                    className="absent"
                    style={{ width: `${(row.absent / maxDeptTotal) * 100}%` }}
                  />
                </div>
                <strong>{row.total}</strong>
              </div>
            ))}
            {data.deptRows.length === 0 ? <p className="empty-copy">ยังไม่มีข้อมูล report</p> : null}
          </div>
        </div>

        <div className="panel report-card center">
          <h3>การเข้างานทั้งโรงงาน</h3>
          <div
            className="report-donut"
            style={{
              background: `conic-gradient(#58991f 0 ${presentPercent}%, #f4a21d ${presentPercent}% ${presentPercent + latePercent}%, #cc1f1f ${presentPercent + latePercent}% 100%)`,
            }}
          >
            <div>
              <strong>{data.totalEmployees}</strong>
              <span>ทั้งหมด</span>
            </div>
          </div>
          <div className="report-legend">
            <LegendRow color="green" label="Present" value={String(data.present)} percent={`${presentPercent.toFixed(1)}%`} />
            <LegendRow color="amber" label="Late" value={String(data.late)} percent={`${latePercent.toFixed(1)}%`} />
            <LegendRow color="red" label="Absent" value={String(data.absent)} percent={`${absentPercent.toFixed(1)}%`} />
          </div>
        </div>

        <div className="panel report-card">
          <h3>Count of attendance_status by dept</h3>
          <div className="mini-pie">
            <div />
          </div>
          <div className="mini-pie-legend">
            {data.deptRows.slice(0, 3).map((row) => (
              <span key={row.dept}>{row.dept}: {row.total}</span>
            ))}
          </div>
        </div>
      </section>

      <section className="panel report-table-panel">
        <table className="table">
          <thead>
            <tr>
              <th>หน่วยงาน</th>
              <th>name</th>
              <th>shift</th>
              <th>shift_start</th>
              <th>scan_in</th>
              <th>Status</th>
              <th>Minutes Late</th>
            </tr>
          </thead>
          <tbody>
            {data.lateRows.map((row) => (
              <tr key={`${row.empId}-${row.scanIn}`}>
                <td>{row.dept}</td>
                <td>{row.name}</td>
                <td>{row.shift}</td>
                <td>{row.shiftStart}</td>
                <td>{row.scanIn}</td>
                <td>{row.status}</td>
                <td>{row.minutesLate}</td>
              </tr>
            ))}
            {data.lateRows.length === 0 ? (
              <tr>
                <td colSpan={7}>กด Load Uploaded Data เพื่อสร้างรายงานจากไฟล์ที่อัปโหลด</td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </section>
    </section>
  );
}

function ReportMetric({
  label,
  tone,
  value,
}: {
  label: string;
  tone?: "green" | "amber" | "red";
  value: number | string;
}) {
  return (
    <div className={`report-metric ${tone ?? ""}`}>
      <strong>{value}</strong>
      <span>{label}</span>
    </div>
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
