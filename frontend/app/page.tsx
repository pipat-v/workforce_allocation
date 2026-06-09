"use client";

import { useEffect, useMemo, useState } from "react";
import {
  AlertTriangle,
  BarChart3,
  BriefcaseBusiness,
  CalendarDays,
  CheckCircle2,
  ChevronDown,
  ClipboardCheck,
  Clock,
  Database,
  Download,
  FileSpreadsheet,
  Home as HomeIcon,
  LayoutGrid,
  LogOut,
  Settings,
  TrendingUp,
  UploadCloud,
  UserX,
  UsersRound,
  X,
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
  original_filename: string | null;
  record_count: number | null;
  created_at: string;
};

type MasterFile = {
  id: string;
  file_type: MasterFileKey;
  file_path: string;
  original_filename: string | null;
  created_at: string;
  is_active: boolean;
};

type DayoffShiftEditorRow = {
  id: string;
  empId: string;
  name: string;
  dayoff: string;
  shift: string;
  raw: Record<string, unknown>;
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
  isoTargetDate: string;
  targetMonthKey: string;
  totalEmployees: number;
  present: number;
  late: number;
  absent: number;
  deptRows: Array<{ dept: string; present: number; late: number; absent: number; total: number }>;
  lateRows: AttendanceRecord[];
  records: AttendanceRecord[];
  timestampRows: AttendanceRecord[];
  monthlyLateCounts: Record<string, number>;
};

const masterFileTypes = [
  { key: "employee_master", label: "รายชื่อพนักงาน" },
  { key: "manpower_plan", label: "Manpower Plan" },
  { key: "skill_matrix", label: "Skill Matrix" },
  { key: "dayoff_shift", label: "Dayoff & Shift" },
] as const;

type TabId =
  | "dashboard"
  | "timestamp"
  | "results"
  | "timestamp_dept"
  | "master"
  | "skill"
  | "report"
  | "setting";

const navItems = [
  { id: "dashboard", label: "Dashboard", icon: HomeIcon },
  { id: "timestamp", label: "Upload Timestamp", icon: UploadCloud },
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
type SortDirection = "asc" | "desc";
type SortState = { key: string; direction: SortDirection } | null;
type AttendanceSortKey =
  | "empId"
  | "name"
  | "dept"
  | "position"
  | "shift"
  | "shiftStart"
  | "scanIn"
  | "status"
  | "minutesLate"
  | "monthlyLate";

const publicWorkspace = "public";

export default function Home() {
  const [activeTab, setActiveTab] = useState<TabId>("dashboard");
  const [masterUploads, setMasterUploads] = useState<MasterUploadState>({
    employee_master: null,
    manpower_plan: null,
    skill_matrix: null,
    dayoff_shift: null,
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
  const [resultsQuery, setResultsQuery] = useState("");
  const [resultsDept, setResultsDept] = useState("all");
  const [resultsStatus, setResultsStatus] = useState("all");
  const [timestampQuery, setTimestampQuery] = useState("");
  const [timestampDept, setTimestampDept] = useState("all");
  const [timestampStatus, setTimestampStatus] = useState("all");
  const [reportLateQuery, setReportLateQuery] = useState("");
  const [reportLateDept, setReportLateDept] = useState("all");
  const [selectedReportDept, setSelectedReportDept] = useState("all");
  const [masterFileHistory, setMasterFileHistory] = useState<MasterFile[]>([]);
  const [dashboardDeptFilter, setDashboardDeptFilter] = useState("all");
  const [warnedIds, setWarnedIds] = useState<Set<string>>(new Set());
  const [warnPending, setWarnPending] = useState<Set<string>>(new Set());
  const [warnCountMap, setWarnCountMap] = useState<Record<string, number>>({});

  const isoTargetDate = reportData?.isoTargetDate ?? "";

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
          activeMasterMap.dayoff_shift?.file_path,
          latestRun.scan_file_path,
        ].filter(Boolean).join("|")
      : "";

  const allDeptOptions = useMemo(
    () => Array.from(new Set((reportData?.records ?? []).map((r) => r.dept))).sort(),
    [reportData],
  );

  const dashboardReport = useMemo<ReportData | null>(() => {
    if (!reportData) return null;
    if (dashboardDeptFilter === "all") return reportData;
    const filtered = reportData.records.filter((r) => r.dept === dashboardDeptFilter);
    const lateFiltered = filtered.filter((r) => r.status === "Late");
    return {
      ...reportData,
      records: filtered,
      timestampRows: filtered,
      totalEmployees: filtered.length,
      present: filtered.filter((r) => r.status === "Present").length,
      late: lateFiltered.length,
      absent: filtered.filter((r) => r.status === "Absent").length,
      lateRows: [...lateFiltered].sort((a, b) => b.minutesLate - a.minutesLate),
      deptRows: reportData.deptRows.filter((r) => r.dept === dashboardDeptFilter),
    };
  }, [reportData, dashboardDeptFilter]);

  const totalEmployees = dashboardReport?.totalEmployees ?? 0;
  const presentPeople = dashboardReport?.present ?? 0;
  const latePeople = dashboardReport?.late ?? 0;
  const absentPeople = dashboardReport?.absent ?? 0;
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

  useEffect(() => {
    if (!isoTargetDate) return;
    supabase
      .from("employee_warnings")
      .select("emp_id")
      .eq("warn_date", isoTargetDate)
      .then(({ data }) => {
        if (data) setWarnedIds(new Set(data.map((r: { emp_id: string }) => r.emp_id)));
      });
  }, [isoTargetDate]);

  useEffect(() => {
    void loadWarnCountMap();
  }, [reportData?.targetMonthKey]);

  async function loadDashboard() {
    await Promise.all([loadRuns(), loadActiveMasters()]);
  }

  async function loadActiveMasters() {
    const { data, error: loadError } = await supabase
      .from("master_data_files")
      .select("id,file_type,file_path,original_filename,is_active,created_at")
      .order("created_at", { ascending: false });

    if (loadError) {
      setError(loadError.message);
      return;
    }

    const allFiles = (data ?? []) as MasterFile[];
    setMasterFileHistory(allFiles);

    const latestByType = new Map<MasterFileKey, MasterFile>();
    for (const item of allFiles) {
      if (item.is_active && !latestByType.has(item.file_type)) {
        latestByType.set(item.file_type, item);
      }
    }

    setActiveMasters(Array.from(latestByType.values()));
  }

  async function loadRuns() {
    const { data, error: loadError } = await supabase
      .from("allocation_runs")
      .select("id,target_date,status,scan_file_path,solver_status,original_filename,record_count,created_at")
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
      dayoff_shift: null,
    });
    setMessage("บันทึก master files แล้ว");
    setIsSavingMasters(false);
    await loadActiveMasters();
  }

  async function saveDayoffShiftRows(rows: DayoffShiftEditorRow[]) {
    setError("");
    setMessage("");

    const fileId = crypto.randomUUID();
    const path = `${publicWorkspace}/masters/dayoff_shift/${fileId}.xlsx`;
    const workbookRows = rows.map((row) => ({
      ...row.raw,
      "User ID (Job Information)": row.empId,
      "ชื่อ นามสกุล": row.name,
      "วันหยุด\nประจำสัปดาห์": row.dayoff,
      "อยู่กะไหน": row.shift,
    }));
    const worksheet = XLSX.utils.json_to_sheet(workbookRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Dayoffandshift");
    const output = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const file = new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const { error: uploadError } = await supabase.storage
      .from("workforce-inputs")
      .upload(path, file, { upsert: true });

    if (uploadError) {
      setError(uploadError.message);
      return;
    }

    const { error: deactivateError } = await supabase
      .from("master_data_files")
      .update({ is_active: false })
      .is("owner_id", null)
      .eq("file_type", "dayoff_shift");

    if (deactivateError) {
      setError(deactivateError.message);
      return;
    }

    const { error: insertError } = await supabase
      .from("master_data_files")
      .insert({
        owner_id: null,
        file_type: "dayoff_shift",
        file_path: path,
        original_filename: "Dayoffandshift-edited.xlsx",
        is_active: true,
      });

    if (insertError) {
      setError(insertError.message);
      return;
    }

    setMessage("บันทึก Dayoff & Shift master แล้ว");
    await loadActiveMasters();
  }

  async function downloadTimestampFile(run: AllocationRun) {
    if (!run.scan_file_path) return;
    const { data, error: urlError } = await supabase.storage
      .from("workforce-inputs")
      .createSignedUrl(run.scan_file_path, 60);
    if (urlError || !data) {
      setError(urlError?.message ?? "ดาวน์โหลดไม่สำเร็จ");
      return;
    }
    window.open(data.signedUrl, "_blank");
  }

  async function deleteRun(run: AllocationRun) {
    const label = run.original_filename ?? run.scan_file_path?.split("/").pop() ?? run.id;
    if (!window.confirm(`ต้องการลบไฟล์ "${label}" ใช่ไหม?\nการลบไม่สามารถย้อนกลับได้`)) return;
    const { error: deleteError } = await supabase
      .from("allocation_runs")
      .delete()
      .eq("id", run.id);
    if (deleteError) {
      setError(deleteError.message);
      return;
    }
    await loadRuns();
  }

  async function deleteMasterFile(file: MasterFile) {
    const label = file.original_filename ?? file.file_type;
    if (!window.confirm(`ต้องการลบ "${label}" ใช่ไหม?\nการลบไม่สามารถย้อนกลับได้`)) return;
    setError("");
    const { error: storageError } = await supabase.storage
      .from("workforce-inputs")
      .remove([file.file_path]);
    if (storageError) {
      setError(storageError.message);
      return;
    }
    const { error: deleteError } = await supabase
      .from("master_data_files")
      .delete()
      .eq("id", file.id);
    if (deleteError) {
      setError(deleteError.message);
      return;
    }
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

    let recordCount: number | null = null;
    try {
      const buffer = await timestampFile.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      recordCount = XLSX.utils.sheet_to_json(firstSheet).length;
    } catch {
      // ignore count error
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
      original_filename: timestampFile.name,
      record_count: recordCount,
      master_file_path: activeMasterMap.employee_master?.file_path,
      manpower_file_path: activeMasterMap.manpower_plan?.file_path,
      skill_file_path: activeMasterMap.skill_matrix?.file_path,
      dayoff_shift_file_path: activeMasterMap.dayoff_shift?.file_path,
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

      const [employeeRows, scanRows, manpowerRows, dayoffShiftRows] = await Promise.all([
        downloadSheetRows(employeeMaster.file_path),
        downloadSheetRows(latestRun.scan_file_path),
        activeMasterMap.manpower_plan
          ? downloadSheetRows(activeMasterMap.manpower_plan.file_path)
          : Promise.resolve([]),
        activeMasterMap.dayoff_shift
          ? downloadSheetRows(activeMasterMap.dayoff_shift.file_path)
          : Promise.resolve([]),
      ]);

      const latestReport = buildReportData(employeeRows, scanRows, manpowerRows, dayoffShiftRows);
      try {
        await saveTimestampWithDeptRows(latestRun.id, latestReport);
      } catch (saveError) {
        const errMsg =
          saveError instanceof Error
            ? saveError.message
            : (saveError as { message?: string })?.message ?? JSON.stringify(saveError);
        setError(`โหลด report ได้ แต่บันทึก Timestamp With Dept ไม่สำเร็จ: ${errMsg}`);
      }
      const scanPaths = Array.from(
        new Set(runs.map((run) => run.scan_file_path).filter(Boolean) as string[]),
      );
      const monthlyLateCounts: Record<string, number> = {};
      const monthlyScanRows = await Promise.all(
        scanPaths.map(async (path) => {
          try {
            return await downloadSheetRows(path);
          } catch {
            return [];
          }
        }),
      );

      for (const rows of monthlyScanRows) {
        const dayReport = buildReportData(employeeRows, rows, manpowerRows, dayoffShiftRows);
        if (dayReport.targetMonthKey !== latestReport.targetMonthKey) continue;

        for (const lateRow of dayReport.lateRows) {
          monthlyLateCounts[lateRow.empId] = (monthlyLateCounts[lateRow.empId] ?? 0) + 1;
        }
      }

      setReportData({
        ...latestReport,
        monthlyLateCounts,
      });
      setLoadedReportKey([
        employeeMaster.file_path,
        activeMasterMap.manpower_plan?.file_path,
        activeMasterMap.dayoff_shift?.file_path,
        latestRun.scan_file_path,
      ].filter(Boolean).join("|"));
    } catch (reportError) {
      setError(reportError instanceof Error ? reportError.message : "โหลด report ไม่สำเร็จ");
    } finally {
      setIsLoadingReport(false);
    }
  }

  async function loadWarnCountMap() {
    const monthKey = reportData?.targetMonthKey;
    if (!monthKey) return;
    const [year, month] = monthKey.split("-");
    const startDate = `${year}-${month}-01`;
    const lastDay = new Date(Number(year), Number(month), 0).getDate();
    const endDate = `${year}-${month}-${String(lastDay).padStart(2, "0")}`;
    const { data: rows, error: queryError } = await supabase
      .from("employee_warnings")
      .select("emp_id, warn_date")
      .gte("warn_date", startDate)
      .lte("warn_date", endDate);
    if (queryError) {
      setError(`โหลดจำนวนเตือนไม่สำเร็จ: ${queryError.message}`);
      return;
    }
    if (!rows) return;
    const map: Record<string, number> = {};
    for (const r of rows as Array<{ emp_id: string; warn_date: string }>) {
      map[r.emp_id] = (map[r.emp_id] ?? 0) + 1;
    }
    setWarnCountMap(map);
  }

  async function toggleWarning(empId: string) {
    if (!isoTargetDate || warnPending.has(empId)) return;
    setWarnPending((s) => new Set(s).add(empId));
    const isWarned = warnedIds.has(empId);
    if (isWarned) {
      const { error: deleteError } = await supabase
        .from("employee_warnings")
        .delete()
        .eq("emp_id", empId)
        .eq("warn_date", isoTargetDate);
      if (deleteError) {
        setError(`ยกเลิกการเตือนไม่สำเร็จ: ${deleteError.message}`);
        setWarnPending((s) => { const n = new Set(s); n.delete(empId); return n; });
        return;
      }
      setWarnedIds((s) => { const n = new Set(s); n.delete(empId); return n; });
    } else {
      const { error: upsertError } = await supabase
        .from("employee_warnings")
        .upsert({ emp_id: empId, warn_date: isoTargetDate }, { onConflict: "emp_id,warn_date" });
      if (upsertError) {
        setError(`บันทึกการเตือนไม่สำเร็จ: ${upsertError.message}`);
        setWarnPending((s) => { const n = new Set(s); n.delete(empId); return n; });
        return;
      }
      setWarnedIds((s) => new Set(s).add(empId));
    }
    setWarnPending((s) => { const n = new Set(s); n.delete(empId); return n; });
    void loadWarnCountMap();
  }

  async function saveTimestampWithDeptRows(runId: string, report: ReportData) {
    const { error: deleteError } = await supabase
      .from("timestamp_with_dept")
      .delete()
      .eq("run_id", runId);

    if (deleteError) {
      throw deleteError;
    }

    const rowMap = new Map<string, object>();
    for (const row of report.timestampRows) {
      rowMap.set(row.empId, {
        run_id: runId,
        target_date: report.targetDate,
        emp_id: row.empId,
        name: row.name,
        dept: row.dept,
        position: row.position,
        shift: row.shift,
        shift_start: row.shiftStart,
        scan_in: row.scanIn,
        attendance_status: row.status,
        minutes_late: row.minutesLate,
      });
    }
    const rows = Array.from(rowMap.values());

    for (let index = 0; index < rows.length; index += 500) {
      const { error: insertError } = await supabase
        .from("timestamp_with_dept")
        .upsert(rows.slice(index, index + 500), { onConflict: "run_id,emp_id" });

      if (insertError) {
        throw insertError;
      }
    }
  }

  return (
    <main className="app-shell">
      <aside className="sidebar">
        <div className="cpf-logo">
          <img alt="WAS" src="/was-logo.png" />
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
          <div className="dashboard-head-left">
            <h2>{activeNav?.label ?? "Dashboard"}</h2>
            {activeTab === "dashboard" && allDeptOptions.length > 0 ? (
              <select
                aria-label="กรองหน่วยงาน"
                className="dept-filter-select"
                value={dashboardDeptFilter}
                onChange={(e) => setDashboardDeptFilter(e.target.value)}
              >
                <option value="all">ทุกหน่วยงาน</option>
                {allDeptOptions.map((dept) => (
                  <option key={dept} value={dept}>{dept}</option>
                ))}
              </select>
            ) : null}
          </div>
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
              <DonutKpiCard
                present={presentPeople}
                late={latePeople}
                absent={absentPeople}
                total={totalEmployees}
                totalActive={totalActivePeople}
              />
            </section>

            <DashboardPanels
              activeMasterMap={activeMasterMap}
              assignedPeople={presentPeople}
              dashboardDeptFilter={dashboardDeptFilter}
              isoTargetDate={isoTargetDate}
              reportData={dashboardReport}
              toggleWarning={toggleWarning}
              totalActivePeople={totalActivePeople}
              warnedIds={warnedIds}
              warnPending={warnPending}
            />
          </>
        ) : null}

        {activeTab === "master" ? (
          <MasterDataPage
            activeMasterMap={activeMasterMap}
            canSaveMasters={canSaveMasters}
            isSavingMasters={isSavingMasters}
            masterFileHistory={masterFileHistory}
            masterUploads={masterUploads}
            onDeleteMasterFile={deleteMasterFile}
            saveDayoffShiftRows={saveDayoffShiftRows}
            saveMasterFiles={saveMasterFiles}
            setMasterUploads={setMasterUploads}
          />
        ) : null}

        {activeTab === "timestamp" ? (
          <TimestampPage
            createDailyRun={createDailyRun}
            deleteRun={deleteRun}
            downloadTimestampFile={downloadTimestampFile}
            hasAllActiveMasters={hasAllActiveMasters}
            isCreatingRun={isCreatingRun}
            latestRun={latestRun}
            runs={runs}
            setTimestampFile={setTimestampFile}
            timestampFile={timestampFile}
          />
        ) : null}

        {activeTab === "results" ? (
          <ResultsPanel
            deptFilter={resultsDept}
            query={resultsQuery}
            reportData={reportData}
            setDeptFilter={setResultsDept}
            setQuery={setResultsQuery}
            setStatusFilter={setResultsStatus}
            standalone
            statusFilter={resultsStatus}
          />
        ) : null}

        {activeTab === "timestamp_dept" ? (
          <TimestampWithDeptPage
            deptFilter={timestampDept}
            query={timestampQuery}
            reportData={reportData}
            setDeptFilter={setTimestampDept}
            setQuery={setTimestampQuery}
            setStatusFilter={setTimestampStatus}
            statusFilter={timestampStatus}
          />
        ) : null}

        {activeTab === "report" ? (
          <ReportDashboard
            deptFilter={reportLateDept}
            isLoadingReport={isLoadingReport}
            loadReportDashboard={loadReportDashboard}
            query={reportLateQuery}
            reportData={reportData}
            selectedDept={selectedReportDept}
            setDeptFilter={setReportLateDept}
            setQuery={setReportLateQuery}
            setSelectedDept={setSelectedReportDept}
            warnCountMap={warnCountMap}
          />
        ) : null}

        {!["dashboard", "master", "timestamp", "results", "timestamp_dept", "report"].includes(activeTab) ? (
          <section className="panel empty-page">
            <h3>{activeNav?.label}</h3>
            <p>แท็บนี้จะเชื่อมข้อมูลจริงในขั้นถัดไป</p>
          </section>
        ) : null}
      </section>
    </main>
  );
}

async function downloadMasterFile(filePath: string, filename: string) {
  const { data, error } = await supabase.storage
    .from("workforce-inputs")
    .createSignedUrl(filePath, 120);
  if (error || !data?.signedUrl) return;
  const a = document.createElement("a");
  a.href = data.signedUrl;
  a.download = filename;
  a.click();
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
    raw: true,
  });

  if (rows.some((row) => "Timestamp" in row && "Employee ID" in row)) {
    return rows;
  }

  const rawRows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, {
    header: 1,
    defval: "",
    raw: true,
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
  dayoffShiftRows: Record<string, unknown>[] = [],
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
  const dayoffShiftMap = buildDayoffShiftMap(dayoffShiftRows);
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

  const latestTimestamp = Array.from(scanByEmp.values())
    .flatMap((entry) => entry.times)
    .sort((a, b) => b.getTime() - a.getTime())[0];
  const targetDate = latestTimestamp?.toLocaleDateString("th-TH") ?? "-";
  const isoTargetDate = latestTimestamp
    ? `${latestTimestamp.getFullYear()}-${String(latestTimestamp.getMonth() + 1).padStart(2, "0")}-${String(latestTimestamp.getDate()).padStart(2, "0")}`
    : "";
  const targetMonthKey = latestTimestamp
    ? `${latestTimestamp.getFullYear()}-${String(latestTimestamp.getMonth() + 1).padStart(2, "0")}`
    : "";

  const baseEmployees = employees.length
    ? employees
    : Array.from(scanByEmp.entries()).map(([empId, entry]) => ({
        empId,
        name: entry.name || empId,
        dept: "ไม่ระบุ",
        position: "พนักงาน",
      }));

  const records: AttendanceRecord[] = baseEmployees.flatMap((employee) => {
    const dayoffShift = dayoffShiftMap.get(employee.empId);
    const scans = scanByEmp.get(employee.empId)?.times ?? [];
    const scanIn = scans.sort((a, b) => a.getTime() - b.getTime())[0];
    const shift = normalizeShiftLabel(dayoffShift?.shift) || "กะ 1";
    const shiftStart =
      deptShiftStart.get(makeDeptShiftKey(employee.dept, shift)) ??
      deptShiftStart.get(makeDeptShiftKey(employee.dept, "")) ??
      "07:00";
    const isScheduledOff = latestTimestamp
      ? isEmployeeDayOff(dayoffShift?.dayoff, latestTimestamp)
      : false;
    if (!scanIn && isScheduledOff) return [];

    const minutesLate = scanIn ? Math.max(0, minutesBetween(shiftStart, scanIn)) : 0;
    const status = !scanIn ? "Absent" : minutesLate > 5 ? "Late" : "Present";

    return [{
      empId: employee.empId,
      name: employee.name,
      dept: employee.dept,
      position: employee.position,
      shift,
      shiftStart,
      scanIn: scanIn ? toTimeText(scanIn) : "-",
      status,
      minutesLate,
    }];
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
    isoTargetDate,
    targetMonthKey,
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
    timestampRows: records,
    monthlyLateCounts: {},
  };
}

function buildDeptShiftStart(rows: Record<string, unknown>[]) {
  const map = new Map<string, string>();
  for (const row of rows) {
    const dept = String(row["หน่วยงาน"] ?? row["dept"] ?? "").trim();
    const shift = normalizeShiftLabel(row["กะ"] ?? row["shift"] ?? row["อยู่กะไหน"]);
    const shiftStart = normalizeTimeText(row["เวลาเข้า"] ?? row["shift_start"]);
    if (!dept || !shiftStart) continue;

    const defaultKey = makeDeptShiftKey(dept, "");
    if (!map.has(defaultKey)) {
      map.set(defaultKey, shiftStart);
    }

    if (shift) {
      const shiftKey = makeDeptShiftKey(dept, shift);
      if (!map.has(shiftKey)) {
        map.set(shiftKey, shiftStart);
      }
    }
  }
  return map;
}

function findRowCol(row: Record<string, unknown>, ...targets: string[]): string {
  const norm = (s: string) => s.replace(/[\s\r\n]+/g, "").toLowerCase();
  const normedTargets = targets.map(norm);
  for (const [key, val] of Object.entries(row)) {
    const k = norm(key);
    if (normedTargets.some((t) => k === t)) return String(val ?? "").trim();
  }
  return "";
}

function buildDayoffShiftMap(rows: Record<string, unknown>[]) {
  const map = new Map<string, { dayoff: string; shift: string }>();
  for (const row of rows) {
    const empId = cleanEmpId(
      row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"]
    );
    if (!empId) continue;
    map.set(empId, {
      dayoff: findRowCol(row, "วันหยุดประจำสัปดาห์", "วันหยุด", "dayoff", "Dayoff", "Day Off"),
      shift: findRowCol(row, "อยู่กะไหน", "shift", "กะ", "Shift"),
    });
  }
  return map;
}

function toDayoffShiftEditorRow(row: Record<string, unknown>, index: number): DayoffShiftEditorRow {
  const empId = cleanEmpId(row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"]);
  const firstName = String(row["First Name (Local)"] ?? "").trim();
  const lastName = String(row["Last Name (Local)"] ?? "").trim();
  const fallbackName = String(row["ชื่อ นามสกุล"] ?? row["Employee Name"] ?? row["Name"] ?? "").trim();
  return {
    id: `${empId || "row"}-${index}`,
    empId,
    name: `${firstName} ${lastName}`.trim() || fallbackName || empId,
    dayoff: findRowCol(row, "วันหยุดประจำสัปดาห์", "วันหยุด", "dayoff", "Dayoff", "Day Off"),
    shift: findRowCol(row, "อยู่กะไหน", "shift", "กะ", "Shift"),
    raw: row,
  };
}

function makeDeptShiftKey(dept: string, shift: string) {
  return `${dept.trim()}|${normalizeShiftKey(shift)}`;
}

function normalizeShiftLabel(value: unknown) {
  const text = String(value ?? "").trim();
  if (!text) return "";

  const match = text.match(/กะ\s*(\d+)/);
  return match ? `กะ ${match[1]}` : text;
}

function normalizeShiftKey(value: unknown) {
  return normalizeShiftLabel(value).replace(/\s+/g, "");
}

function isEmployeeDayOff(dayoff: string | undefined, targetDate: Date) {
  const value = String(dayoff ?? "").trim();
  if (!value) return false;
  if (value === "พระ") return isBuddhistHolyDay(targetDate);
  return value === getThaiWeekdayCode(targetDate);
}

function getThaiWeekdayCode(date: Date) {
  return ["อา", "จ", "อ", "พ", "พฤ", "ศ", "ส"][date.getDay()];
}

const buddhistHolyDaysByYear: Record<string, Set<string>> = {
  "2025": new Set([
    "2025-01-05", "2025-01-13", "2025-01-20", "2025-01-29",
    "2025-02-05", "2025-02-12", "2025-02-20", "2025-02-27",
    "2025-03-06", "2025-03-14", "2025-03-22", "2025-03-29",
    "2025-04-05", "2025-04-13", "2025-04-21", "2025-04-28",
    "2025-05-04", "2025-05-12", "2025-05-20", "2025-05-27",
    "2025-06-03", "2025-06-11", "2025-06-18", "2025-06-25",
    "2025-07-04", "2025-07-10", "2025-07-18", "2025-07-24",
    "2025-08-02", "2025-08-09", "2025-08-17", "2025-08-23",
    "2025-09-01", "2025-09-07", "2025-09-15", "2025-09-22",
    "2025-10-01", "2025-10-07", "2025-10-14", "2025-10-21", "2025-10-29",
    "2025-11-05", "2025-11-12", "2025-11-20", "2025-11-27",
    "2025-12-05", "2025-12-11", "2025-12-19", "2025-12-26",
  ]),
  "2026": new Set([
    "2026-01-03", "2026-01-11", "2026-01-18", "2026-01-26",
    "2026-02-02", "2026-02-10", "2026-02-16", "2026-02-24",
    "2026-03-03", "2026-03-11", "2026-03-18", "2026-03-26",
    "2026-04-02", "2026-04-10", "2026-04-16", "2026-04-24",
    "2026-05-01", "2026-05-09", "2026-05-16", "2026-05-24", "2026-05-31",
    "2026-06-08", "2026-06-14", "2026-06-22", "2026-06-29",
    "2026-07-07", "2026-07-14", "2026-07-22", "2026-07-29", "2026-07-30",
    "2026-08-06", "2026-08-13", "2026-08-21", "2026-08-28",
    "2026-09-05", "2026-09-11", "2026-09-19", "2026-09-26",
    "2026-10-04", "2026-10-11", "2026-10-19", "2026-10-26",
    "2026-11-03", "2026-11-09", "2026-11-17", "2026-11-24",
    "2026-12-02", "2026-12-09", "2026-12-17", "2026-12-24",
  ]),
  "2027": new Set([
    "2027-01-01", "2027-01-08", "2027-01-15", "2027-01-22", "2027-01-30",
    "2027-02-06", "2027-02-13", "2027-02-21", "2027-02-28",
    "2027-03-07", "2027-03-14", "2027-03-22", "2027-03-29",
    "2027-04-06", "2027-04-13", "2027-04-21", "2027-04-28",
    "2027-05-06", "2027-05-12", "2027-05-20", "2027-05-27",
    "2027-06-03", "2027-06-11", "2027-06-18", "2027-06-26",
    "2027-07-02", "2027-07-10", "2027-07-17", "2027-07-25",
    "2027-08-01", "2027-08-09", "2027-08-15", "2027-08-23", "2027-08-30",
    "2027-09-07", "2027-09-14", "2027-09-22", "2027-09-29",
    "2027-10-06", "2027-10-13", "2027-10-21", "2027-10-28",
    "2027-11-05", "2027-11-11", "2027-11-19", "2027-11-26",
    "2027-12-05", "2027-12-10", "2027-12-19", "2027-12-25",
  ]),
};

function isBuddhistHolyDay(date: Date) {
  const yearKey = String(date.getFullYear());
  return buddhistHolyDaysByYear[yearKey]?.has(toDateKey(date)) ?? false;
}

function toDateKey(date: Date) {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
  ].join("-");
}

function cleanEmpId(value: unknown) {
  return String(value ?? "")
    .replace(/\u00a0/g, "")
    .replace(/\s+/g, "")
    .replace(/\.0$/, "");
}

function getAttendanceSortValue(
  row: AttendanceRecord,
  key: string,
  monthlyLateCounts: Record<string, number> = {},
) {
  if (key === "minutesLate") return row.minutesLate;
  if (key === "monthlyLate") return monthlyLateCounts[row.empId] ?? 0;
  return (row as Record<string, unknown>)[key] as string ?? "";
}

function sortAttendanceRows(
  rows: AttendanceRecord[],
  sort: SortState,
  monthlyLateCounts: Record<string, number> = {},
) {
  if (!sort) return rows;

  return [...rows].sort((a, b) => {
    const aValue = getAttendanceSortValue(a, sort.key, monthlyLateCounts);
    const bValue = getAttendanceSortValue(b, sort.key, monthlyLateCounts);

    if (typeof aValue === "number" && typeof bValue === "number") {
      return sort.direction === "asc" ? aValue - bValue : bValue - aValue;
    }

    const comparison = String(aValue).localeCompare(String(bValue), "th", {
      numeric: true,
      sensitivity: "base",
    });
    return sort.direction === "asc" ? comparison : -comparison;
  });
}

function SortButton({
  children,
  columnKey,
  setSort,
  sort,
}: {
  children: ReactNode;
  columnKey: string;
  setSort?: (sort: SortState) => void;
  sort?: SortState;
}) {
  const active = sort?.key === columnKey;
  return (
    <button
      className={`sort-button ${active ? "active" : ""}`}
      disabled={!setSort}
      onClick={() => {
        if (!setSort) return;
        if (!sort || sort.key !== columnKey) {
          setSort({ key: columnKey, direction: "asc" });
        } else {
          setSort({ key: columnKey, direction: sort.direction === "asc" ? "desc" : "asc" });
        }
      }}
      type="button"
    >
      {children}
      <span>{active ? (sort?.direction === "asc" ? "▲" : "▼") : "↕"}</span>
    </button>
  );
}

function parseTimestamp(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return excelSerialDateToDate(value);
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

  if (typeof value === "number" && Number.isFinite(value)) {
    const totalMinutes = Math.round((value % 1) * 24 * 60);
    const hours = Math.floor(totalMinutes / 60) % 24;
    const minutes = totalMinutes % 60;
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`;
  }

  const text = String(value ?? "").trim();
  const match = text.match(/(\d{1,2}):(\d{2})/);
  if (!match) return "";
  return `${match[1].padStart(2, "0")}:${match[2]}`;
}

function excelSerialDateToDate(value: number) {
  const excelEpoch = Date.UTC(1899, 11, 30);
  return new Date(excelEpoch + value * 24 * 60 * 60 * 1000);
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

function formatLateTime(minutes: number): string {
  if (!minutes) return "0 นาที";
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  if (h === 0) return `${m} นาที`;
  if (m === 0) return `${h} ชม.`;
  return `${h} ชม. ${m} นาที`;
}

function getSafeFileExtension(filename: string) {
  const match = filename.toLowerCase().match(/\.(csv|xlsx|xls)$/);
  return match ? `.${match[1]}` : "";
}

function DonutKpiCard({
  present,
  late,
  absent,
  total,
  totalActive,
}: {
  present: number;
  late: number;
  absent: number;
  total: number;
  totalActive: number;
}) {
  const presentPct = total ? (present / total) * 100 : 0;
  const latePct = total ? (late / total) * 100 : 0;
  const absentPct = total ? (absent / total) * 100 : 0;

  return (
    <article className="kpi-card kpi-donut">
      <div className="kpi-bar-chart">
        <span className="kpi-bar-label">{totalActive} คน</span>
        <div className="kpi-stacked-bar">
          <div className="kpi-bar-fill present" style={{ width: `${presentPct}%` }} />
          <div className="kpi-bar-fill late" style={{ width: `${latePct}%` }} />
          <div className="kpi-bar-fill absent" style={{ width: `${absentPct}%` }} />
        </div>
      </div>
      <div className="legend compact">
        <LegendRow color="green" label="Present" value={String(present)} percent={`${presentPct.toFixed(1)}%`} />
        <LegendRow color="amber" label="Late" value={String(late)} percent={`${latePct.toFixed(1)}%`} />
        <LegendRow color="red" label="Absent" value={String(absent)} percent={`${absentPct.toFixed(1)}%`} />
      </div>
    </article>
  );
}

function exportLateAbsentToExcel(
  rows: AttendanceRecord[],
  monthlyLateCounts: Record<string, number>,
  deptLabel: string,
) {
  const exportRows = rows
    .filter((r) => r.status === "Late" || r.status === "Absent")
    .sort((a, b) => {
      if (a.status !== b.status) return a.status === "Absent" ? -1 : 1;
      return b.minutesLate - a.minutesLate;
    })
    .map((r, i) => ({
      "ลำดับ": i + 1,
      "รหัสพนักงาน": r.empId,
      "ชื่อ-สกุล": r.name,
      "หน่วยงาน": r.dept,
      "ตำแหน่ง": r.position,
      "กะ": r.shift,
      "เวลาเข้างาน": r.shiftStart,
      "Scan In": r.scanIn,
      "สถานะ": r.status,
      "สาย (นาที)": r.minutesLate,
      "สายสะสมเดือนนี้ (ครั้ง)": monthlyLateCounts[r.empId] ?? 0,
    }));

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Late-Absent");
  const filename = `late-absent_${deptLabel}_${new Date().toLocaleDateString("th-TH").replace(/\//g, "-")}.xlsx`;
  XLSX.writeFile(wb, filename);
}

function DashboardPanels({
  activeMasterMap,
  assignedPeople,
  dashboardDeptFilter,
  isoTargetDate,
  reportData,
  toggleWarning,
  totalActivePeople,
  warnedIds,
  warnPending,
}: {
  activeMasterMap: Partial<Record<MasterFileKey, MasterFile>>;
  assignedPeople: number;
  dashboardDeptFilter: string;
  isoTargetDate: string;
  reportData: ReportData | null;
  toggleWarning: (empId: string) => Promise<void>;
  totalActivePeople: number;
  warnedIds: Set<string>;
  warnPending: Set<string>;
}) {
  const [detailStatusFilter, setDetailStatusFilter] = useState("all");
  const [detailSort, setDetailSort_] = useState<SortState>(null);
  const setDetailSort = setDetailSort_ as (sort: SortState) => void;

  const total = reportData?.totalEmployees ?? 0;
  const present = reportData?.present ?? 0;
  const late = reportData?.late ?? 0;
  const absent = reportData?.absent ?? 0;
  const presentPct = total ? (present / total) * 100 : 0;
  const latePct = total ? (late / total) * 100 : 0;
  const absentPct = total ? (absent / total) * 100 : 0;
  const topDeptRows = reportData?.deptRows ?? [];
  const dashboardLateRows = reportData?.lateRows ?? [];
  const maxDeptTotal = Math.max(...topDeptRows.map((row) => row.total), 1);
  const monthlyLateCounts = reportData?.monthlyLateCounts ?? {};
  const donutStyle = total
    ? {
        background: `conic-gradient(
          #10b981 0 ${presentPct}%,
          #f59e0b ${presentPct}% ${presentPct + latePct}%,
          #ef4444 ${presentPct + latePct}% 100%
        )`,
      }
    : { background: "#e2e8f0" };

  const allRecords = reportData?.records ?? [];
  const statusOrder: Record<string, number> = { Absent: 0, Late: 1, Present: 2 };
  const baseDetailRows = allRecords.filter((r) => detailStatusFilter === "all" || r.status === detailStatusFilter);
  const detailRows = detailSort
    ? sortAttendanceRows(baseDetailRows, detailSort, monthlyLateCounts)
    : [...baseDetailRows].sort((a, b) => {
        const sd = (statusOrder[a.status] ?? 3) - (statusOrder[b.status] ?? 3);
        if (sd !== 0) return sd;
        return b.minutesLate - a.minutesLate;
      });

  const lateAbsentCount = allRecords.filter((r) => r.status === "Late" || r.status === "Absent").length;

  return (
    <>
      <section className="dashboard-grid">
        <section className="panel dashboard-late-card">
          <div className="panel-title-row">
            <h3>คนที่มาสาย</h3>
            <span className="table-count">{dashboardLateRows.length} คน</span>
          </div>
          <div className="late-preview-table">
            <table className="table compact-table">
              <thead>
                <tr>
                  <th>ชื่อ</th>
                  <th>หน่วยงาน</th>
                  <th>เข้างาน</th>
                  <th>สาย</th>
                  <th>เตือน</th>
                </tr>
              </thead>
              <tbody>
                {dashboardLateRows.map((row) => {
                  const warned = warnedIds.has(row.empId);
                  const pending = warnPending.has(row.empId);
                  return (
                  <tr key={`dashboard-late-${row.empId}-${row.scanIn}`} className={warned ? "row-warned" : ""}>
                    <td>{row.name}</td>
                    <td><span className="dept-chip">{row.dept}</span></td>
                    <td>{row.scanIn}</td>
                    <td><span className="late-minutes-badge">{formatLateTime(row.minutesLate)}</span></td>
                    <td>
                      <button
                        className={`warn-btn${warned ? " warned" : ""}`}
                        disabled={pending || !isoTargetDate}
                        onClick={() => void toggleWarning(row.empId)}
                        title={warned ? "ยกเลิกการเตือน" : "บันทึกว่าเตือนแล้ว"}
                        type="button"
                      >
                        {warned ? "✓ เตือนแล้ว" : "เตือน"}
                      </button>
                    </td>
                  </tr>
                  );
                })}
                {dashboardLateRows.length === 0 ? (
                  <tr><td colSpan={5}>ยังไม่มีข้อมูลคนมาสาย</td></tr>
                ) : null}
              </tbody>
            </table>
          </div>
        </section>

        <section className="panel dept-panel">
          <div className="panel-title-row">
            <h3>พนักงานตามหน่วยงาน</h3>
          </div>
          <div className="dept-bars">
            {topDeptRows.length === 0 ? (
              <p className="empty-copy">ยังไม่มีข้อมูล</p>
            ) : null}
            {topDeptRows.map((row, index) => (
              <div className="dept-row" key={row.dept}>
                <span className="dept-rank">{index + 1}</span>
                <div className="dept-row-content">
                  <div className="dept-row-label">
                    <span>{row.dept}</span>
                    <strong>{row.total}</strong>
                  </div>
                  <div className="bar-track">
                    <div className="bar-fill" style={{ width: `${(row.total / maxDeptTotal) * 100}%` }} />
                  </div>
                </div>
              </div>
            ))}
          </div>
        </section>

      </section>

      {/* Employee detail table */}
      <section className="panel detail-attendance-panel">
        <div className="panel-title-row">
          <h3>
            สถานะพนักงานรายคน
            {dashboardDeptFilter !== "all" ? ` · ${dashboardDeptFilter}` : ""}
          </h3>
          <div className="table-actions">
            <select
              aria-label="กรองสถานะ"
              value={detailStatusFilter}
              onChange={(e) => setDetailStatusFilter(e.target.value)}
            >
              <option value="all">ทุกสถานะ</option>
              <option value="Absent">Absent</option>
              <option value="Late">Late</option>
              <option value="Present">Present</option>
            </select>
            <button
              className="primary-button small"
              disabled={lateAbsentCount === 0}
              onClick={() => exportLateAbsentToExcel(allRecords, monthlyLateCounts, dashboardDeptFilter === "all" ? "ทุกหน่วยงาน" : dashboardDeptFilter)}
              type="button"
            >
              <Download size={14} />
              Export Late/Absent
            </button>
          </div>
        </div>
        <div className="table-scroll">
          <table className="table data-table">
            <thead>
              <tr>
                <th>No.</th>
                <th><SortButton columnKey="name" setSort={setDetailSort} sort={detailSort}>ชื่อ-สกุล</SortButton></th>
                <th><SortButton columnKey="dept" setSort={setDetailSort} sort={detailSort}>หน่วยงาน</SortButton></th>
                <th><SortButton columnKey="shift" setSort={setDetailSort} sort={detailSort}>กะ</SortButton></th>
                <th><SortButton columnKey="shiftStart" setSort={setDetailSort} sort={detailSort}>เวลาเข้างาน</SortButton></th>
                <th><SortButton columnKey="scanIn" setSort={setDetailSort} sort={detailSort}>Scan In</SortButton></th>
                <th><SortButton columnKey="status" setSort={setDetailSort} sort={detailSort}>สถานะ</SortButton></th>
                <th><SortButton columnKey="minutesLate" setSort={setDetailSort} sort={detailSort}>สาย</SortButton></th>
                <th><SortButton columnKey="monthlyLate" setSort={setDetailSort} sort={detailSort}>สายเดือนนี้</SortButton></th>
                <th>เสี่ยง</th>
              </tr>
            </thead>
            <tbody key={detailStatusFilter}>
              {detailRows.map((row, index) => {
                const monthlyLate = monthlyLateCounts[row.empId] ?? 0;
                const isRisk = monthlyLate >= 3;
                return (
                  <tr key={`${row.empId}-${row.scanIn}-detail`} className={isRisk ? "row-risk" : ""}>
                    <td>{index + 1}</td>
                    <td>{row.name}</td>
                    <td>{row.dept}</td>
                    <td>{row.shift}</td>
                    <td>{row.shiftStart}</td>
                    <td>{row.scanIn}</td>
                    <td><span className={`status-pill ${row.status.toLowerCase()}`}>{row.status}</span></td>
                    <td>{row.status !== "Absent" ? formatLateTime(row.minutesLate) : "-"}</td>
                    <td>
                      <span className={monthlyLate >= 3 ? "monthly-late-high" : ""}>
                        {monthlyLate > 0 ? monthlyLate : "-"}
                      </span>
                    </td>
                    <td>{isRisk ? <span className="risk-badge">เสี่ยง</span> : null}</td>
                  </tr>
                );
              })}
              {detailRows.length === 0 ? (
                <tr><td colSpan={10}>ยังไม่มีข้อมูล</td></tr>
              ) : null}
            </tbody>
          </table>
        </div>
        <p className="panel-note">
          <span className="risk-badge" style={{ marginRight: 6 }}>เสี่ยง</span>
          = มาสายสะสม ≥ 3 ครั้งในเดือนนี้ · Export จะรวมเฉพาะ Late และ Absent
        </p>
      </section>

      <ResultsPanel reportData={reportData} />
    </>
  );
}

const masterColumnColors: Record<MasterFileKey, string> = {
  employee_master: "#2563eb",
  manpower_plan: "#d97706",
  skill_matrix: "#10b981",
  dayoff_shift: "#7c3aed",
};

function MasterDataPage({
  activeMasterMap,
  canSaveMasters,
  isSavingMasters,
  masterFileHistory,
  masterUploads,
  onDeleteMasterFile,
  saveDayoffShiftRows,
  saveMasterFiles,
  setMasterUploads,
}: {
  activeMasterMap: Partial<Record<MasterFileKey, MasterFile>>;
  canSaveMasters: boolean;
  isSavingMasters: boolean;
  masterFileHistory: MasterFile[];
  masterUploads: MasterUploadState;
  onDeleteMasterFile: (file: MasterFile) => Promise<void>;
  saveDayoffShiftRows: (rows: DayoffShiftEditorRow[]) => Promise<void>;
  saveMasterFiles: () => Promise<void>;
  setMasterUploads: Dispatch<SetStateAction<MasterUploadState>>;
}) {
  return (
    <section className="md-page">
      <div className="md-columns-bar">
        <div>
          <h3>Master Data</h3>
          <p>อัปโหลดไฟล์หลัก 4 ไฟล์ ระบบจะใช้ชุดล่าสุดกับ daily run อัตโนมัติ</p>
        </div>
        <button
          className="primary-button"
          disabled={!canSaveMasters || isSavingMasters}
          onClick={saveMasterFiles}
          type="button"
        >
          <UploadCloud size={17} />
          {isSavingMasters ? "Saving..." : "Save Master Files"}
        </button>
      </div>

      <div className="md-columns">
        {masterFileTypes.map((item) => {
          const pendingFile = masterUploads[item.key];
          const fileHistory = masterFileHistory.filter((f) => f.file_type === item.key);
          const inputId = `master-input-${item.key}`;
          const color = masterColumnColors[item.key];

          return (
            <div className="md-column panel" key={item.key}>
              <div className="md-column-header" style={{ borderTopColor: color }}>
                <span className="md-column-title">{item.label}</span>
              </div>

              <label className={`ts-dropzone compact-dropzone ${pendingFile ? "has-file" : ""}`}>
                <UploadCloud size={28} />
                {pendingFile ? (
                  <>
                    <strong>{pendingFile.name}</strong>
                    <span>{(pendingFile.size / 1024).toFixed(0)} KB · คลิกเพื่อเปลี่ยนไฟล์</span>
                  </>
                ) : (
                  <>
                    <strong>ลากไฟล์มาวางที่นี่ หรือ คลิกเพื่อเลือกไฟล์</strong>
                    <span>รองรับ .xlsx, .xls, .csv</span>
                  </>
                )}
                <input
                  id={inputId}
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

              <div className="md-col-history">
                <h4>ประวัติการอัปโหลด</h4>
                <div className="ts-history-list">
                  {fileHistory.length === 0 ? (
                    <p className="empty-copy" style={{ padding: "8px 0" }}>ยังไม่มีประวัติ</p>
                  ) : null}
                  {fileHistory.map((file) => {
                    const dateText = new Date(file.created_at).toLocaleString("th-TH", {
                      day: "numeric",
                      month: "numeric",
                      year: "2-digit",
                      hour: "2-digit",
                      minute: "2-digit",
                      hour12: false,
                    });
                    return (
                      <div className="ts-history-row" key={file.id}>
                        <div className="ts-history-info">
                          <div className="ts-history-name-row">
                            <strong>{file.original_filename ?? "ไฟล์"}</strong>
                            {file.is_active ? (
                              <span className="status-pill uploaded">Active</span>
                            ) : null}
                          </div>
                          <span>{dateText}</span>
                        </div>
                        <div className="ts-history-actions">
                          <button
                            className="icon-button"
                            onClick={() => downloadMasterFile(file.file_path, file.original_filename ?? "download.xlsx")}
                            title="ดาวน์โหลด"
                            type="button"
                          >
                            <Download size={15} />
                          </button>
                          <button
                            className="icon-button danger"
                            onClick={() => void onDeleteMasterFile(file)}
                            title="ลบ"
                            type="button"
                          >
                            <X size={15} />
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          );
        })}
      </div>

      <DayoffShiftEditor
        activeFile={activeMasterMap.dayoff_shift}
        saveDayoffShiftRows={saveDayoffShiftRows}
      />
    </section>
  );
}

function DayoffShiftEditor({
  activeFile,
  saveDayoffShiftRows,
}: {
  activeFile?: MasterFile;
  saveDayoffShiftRows: (rows: DayoffShiftEditorRow[]) => Promise<void>;
}) {
  const [rows, setRows] = useState<DayoffShiftEditorRow[]>([]);
  const [query, setQuery] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const dayoffOptions = ["จ", "อ", "พ", "พฤ", "ศ", "ส", "อา", "พระ"];
  const shiftOptions = Array.from(new Set([
    "กะ1",
    "กะ2",
    "กะ3",
    ...rows.map((row) => row.shift).filter(Boolean),
  ]));
  const normalizedQuery = query.trim().toLowerCase();
  const filteredRows = rows.filter((row) => {
    if (!normalizedQuery) return true;
    return [row.empId, row.name, row.dayoff, row.shift]
      .some((value) => value.toLowerCase().includes(normalizedQuery));
  });

  useEffect(() => {
    if (!activeFile?.file_path) {
      setRows([]);
      return;
    }

    let isMounted = true;
    setIsLoading(true);
    downloadSheetRows(activeFile.file_path)
      .then((sourceRows) => {
        if (!isMounted) return;
        setRows(sourceRows.map(toDayoffShiftEditorRow));
      })
      .catch(() => {
        if (!isMounted) return;
        setRows([]);
      })
      .finally(() => {
        if (!isMounted) return;
        setIsLoading(false);
      });

    return () => {
      isMounted = false;
    };
  }, [activeFile?.file_path]);

  function updateRow(id: string, field: "dayoff" | "shift", value: string) {
    setRows((current) =>
      current.map((row) => (
        row.id === id
          ? {
              ...row,
              [field]: value,
              raw: {
                ...row.raw,
                [field === "dayoff" ? "วันหยุด\nประจำสัปดาห์" : "อยู่กะไหน"]: value,
              },
            }
          : row
      )),
    );
  }

  async function handleSave() {
    setIsSaving(true);
    try {
      await saveDayoffShiftRows(rows);
    } finally {
      setIsSaving(false);
    }
  }

  return (
    <section className="panel dayoff-editor-panel">
      <div className="panel-title-row">
        <div>
          <h3>แก้ไข Dayoff & Shift</h3>
          <p>ปรับวันหยุดประจำสัปดาห์และกะทำงานรายคน แล้วบันทึกเป็น master active ชุดใหม่</p>
        </div>
        <button
          className="primary-button"
          disabled={!rows.length || isSaving}
          onClick={handleSave}
          type="button"
        >
          <UploadCloud size={17} />
          {isSaving ? "Saving" : "Save Changes"}
        </button>
      </div>

      <div className="table-filters dayoff-editor-filters">
        <input
          aria-label="ค้นหา dayoff shift"
          placeholder="ค้นหา รหัส ชื่อ วันหยุด กะ"
          type="search"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
        />
        <span>{filteredRows.length.toLocaleString()} / {rows.length.toLocaleString()} คน</span>
      </div>

      <div className="dayoff-editor-table">
        <table className="table">
          <thead>
            <tr>
              <th>Emp ID</th>
              <th>ชื่อ</th>
              <th>Dayoff</th>
              <th>Shift</th>
            </tr>
          </thead>
          <tbody>
            {filteredRows.map((row) => (
              <tr key={row.id}>
                <td>{row.empId}</td>
                <td>{row.name}</td>
                <td>
                  <select
                    value={row.dayoff}
                    onChange={(event) => updateRow(row.id, "dayoff", event.target.value)}
                  >
                    <option value="">-</option>
                    {dayoffOptions.map((option) => (
                      <option key={option} value={option}>{option}</option>
                    ))}
                  </select>
                </td>
                <td>
                  <select
                    value={row.shift}
                    onChange={(event) => updateRow(row.id, "shift", event.target.value)}
                  >
                    <option value="">-</option>
                    {shiftOptions.map((option) => (
                      <option key={option} value={option}>{option}</option>
                    ))}
                  </select>
                </td>
              </tr>
            ))}
            {!activeFile ? (
              <tr>
                <td colSpan={4}>อัปโหลด Dayoff & Shift master ก่อน จึงจะแก้ไขในหน้านี้ได้</td>
              </tr>
            ) : null}
            {activeFile && isLoading ? (
              <tr>
                <td colSpan={4}>Loading Dayoff & Shift...</td>
              </tr>
            ) : null}
            {activeFile && !isLoading && filteredRows.length === 0 ? (
              <tr>
                <td colSpan={4}>ไม่พบข้อมูลที่ค้นหา</td>
              </tr>
            ) : null}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function TimestampPage({
  createDailyRun,
  deleteRun,
  downloadTimestampFile,
  hasAllActiveMasters,
  isCreatingRun,
  latestRun,
  runs,
  setTimestampFile,
  timestampFile,
}: {
  createDailyRun: () => Promise<void>;
  deleteRun: (run: AllocationRun) => Promise<void>;
  downloadTimestampFile: (run: AllocationRun) => Promise<void>;
  hasAllActiveMasters: boolean;
  isCreatingRun: boolean;
  latestRun?: AllocationRun;
  runs: AllocationRun[];
  setTimestampFile: (file: File | null) => void;
  timestampFile: File | null;
}) {
  return (
    <section className="workspace-grid">
      <section className="panel ts-upload-panel">
        <div className="ts-panel-header">
          <h3>Upload Timestamp</h3>
          <p>อัปโหลดไฟล์ timestamp รายวัน ระบบจะใช้ master data ชุด active ล่าสุด</p>
        </div>

        <label className={`ts-dropzone ${timestampFile ? "has-file" : ""}`}>
          <UploadCloud size={38} />
          {timestampFile ? (
            <>
              <strong>{timestampFile.name}</strong>
              <span>{(timestampFile.size / 1024).toFixed(0)} KB · คลิกเพื่อเปลี่ยนไฟล์</span>
            </>
          ) : (
            <>
              <strong>ลากไฟล์มาวางที่นี่</strong>
              <span>หรือคลิกเพื่อเลือกไฟล์ · รองรับ .csv, .xlsx, .xls</span>
            </>
          )}
          <input
            type="file"
            accept=".csv,.xlsx,.xls"
            onChange={(event) => setTimestampFile(event.target.files?.[0] ?? null)}
          />
        </label>

        <div className={`ts-master-status ${hasAllActiveMasters ? "ready" : "not-ready"}`}>
          <span className={`ts-status-dot ${hasAllActiveMasters ? "ready" : "not-ready"}`} />
          {hasAllActiveMasters
            ? "Master files พร้อมใช้งานครบทั้ง 4 ไฟล์"
            : "ต้องมี master files ครบ 4 ไฟล์ก่อนสร้าง daily run"}
        </div>

        <button
          className="primary-button ts-submit-btn"
          disabled={!timestampFile || !hasAllActiveMasters || isCreatingRun}
          onClick={createDailyRun}
          type="button"
        >
          <ClipboardCheck size={18} />
          {isCreatingRun ? "กำลังสร้าง..." : "Create Daily Run"}
        </button>
      </section>

      <section className="panel ts-history-panel">
        {latestRun ? (
          <div className="ts-latest-run">
            <div className="ts-latest-header">
              <span>Latest Run</span>
              <span className={`status-pill ${latestRun.status.toLowerCase()}`}>{latestRun.status}</span>
            </div>
            <strong className="ts-latest-filename">
              {latestRun.original_filename ?? latestRun.scan_file_path?.split("/").pop() ?? "-"}
            </strong>
            <span className="ts-latest-meta">
              {latestRun.record_count != null ? `${latestRun.record_count.toLocaleString()} รายการ · ` : ""}
              {new Date(latestRun.created_at).toLocaleString("th-TH", {
                day: "numeric",
                month: "long",
                year: "numeric",
                hour: "2-digit",
                minute: "2-digit",
                hour12: false,
              })}
            </span>
          </div>
        ) : null}

        <div className="ts-history-header">
          <h3>ประวัติการอัปโหลด</h3>
          <span className="table-count">{runs.length} ไฟล์</span>
        </div>

        <div className="ts-history-list">
          {runs.length === 0 ? (
            <p className="empty-copy">ยังไม่มีประวัติการอัปโหลด</p>
          ) : null}
          {runs.map((run) => {
            const filename = run.original_filename
              ?? run.scan_file_path?.split("/").pop()
              ?? "-";
            const dateText = new Date(run.created_at).toLocaleString("th-TH", {
              day: "numeric",
              month: "numeric",
              year: "2-digit",
              hour: "2-digit",
              minute: "2-digit",
              hour12: false,
            });
            const meta = [
              run.record_count != null ? `${run.record_count.toLocaleString()} รายการ` : null,
              dateText,
            ].filter(Boolean).join(" · ");
            return (
              <div className="ts-history-row" key={run.id}>
                <div className="ts-history-info">
                  <strong>{filename}</strong>
                  <span>{meta}</span>
                </div>
                <div className="ts-history-actions">
                  <button
                    className="icon-button"
                    onClick={() => void downloadTimestampFile(run)}
                    title="ดาวน์โหลด"
                    type="button"
                  >
                    <Download size={15} />
                  </button>
                  <button
                    className="icon-button danger"
                    onClick={() => void deleteRun(run)}
                    title="ลบ"
                    type="button"
                  >
                    <X size={15} />
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      </section>
    </section>
  );
}


const pageSize = 10;

function TimestampWithDeptPage({
  deptFilter,
  query,
  reportData,
  setDeptFilter,
  setQuery,
  setStatusFilter,
  statusFilter,
}: {
  deptFilter: string;
  query: string;
  reportData: ReportData | null;
  setDeptFilter: (value: string) => void;
  setQuery: (value: string) => void;
  setStatusFilter: (value: string) => void;
  statusFilter: string;
}) {
  const [page, setPage] = useState(1);
  const [sort, setSort_] = useState<SortState>(null);
  const setSort = setSort_ as (sort: SortState) => void;

  const sourceRows = reportData?.timestampRows ?? [];
  const deptOptions = Array.from(new Set(sourceRows.map((record) => record.dept))).sort();
  const normalizedQuery = query.trim().toLowerCase();
  const allRows = sourceRows.filter((record) => {
    const matchesQuery = !normalizedQuery || [
      record.empId,
      record.name,
      record.dept,
      record.position,
      record.scanIn,
      record.status,
    ].some((value) => String(value).toLowerCase().includes(normalizedQuery));
    const matchesDept = deptFilter === "all" || record.dept === deptFilter;
    const matchesStatus = statusFilter === "all" || record.status === statusFilter;
    return matchesQuery && matchesDept && matchesStatus;
  });
  const sortedRows = sortAttendanceRows(allRows, sort);
  const totalPages = Math.max(1, Math.ceil(sortedRows.length / pageSize));
  const safePage = Math.min(page, totalPages);
  const rows = sortedRows.slice((safePage - 1) * pageSize, safePage * pageSize);

  function updateFilter(callback: () => void) {
    callback();
    setPage(1);
  }

  return (
    <section className="panel results-panel">
      <div className="panel-title-row">
        <h3>Timestamp With Dept</h3>
        <span className="table-count">{allRows.length} rows</span>
      </div>
      <div className="table-filters">
        <input
          aria-label="ค้นหา timestamp"
          placeholder="ค้นหา รหัส ชื่อ หน่วยงาน สถานะ"
          type="search"
          value={query}
          onChange={(event) => updateFilter(() => setQuery(event.target.value))}
        />
        <select
          aria-label="หน่วยงาน"
          value={deptFilter}
          onChange={(event) => updateFilter(() => setDeptFilter(event.target.value))}
        >
          <option value="all">ทุกหน่วยงาน</option>
          {deptOptions.map((dept) => (
            <option key={dept} value={dept}>{dept}</option>
          ))}
        </select>
        <select
          aria-label="สถานะ"
          value={statusFilter}
          onChange={(event) => updateFilter(() => setStatusFilter(event.target.value))}
        >
          <option value="all">ทุกสถานะ</option>
          <option value="Present">Present</option>
          <option value="Late">Late</option>
          <option value="Absent">Absent</option>
        </select>
        <button
          className="ghost-button"
          onClick={() => updateFilter(() => {
            setQuery("");
            setDeptFilter("all");
            setStatusFilter("all");
          })}
          type="button"
        >
          Clear
        </button>
      </div>
      <div className="table-scroll">
        <table className="table data-table">
          <thead>
            <tr>
              <th><SortButton columnKey="empId" setSort={setSort} sort={sort}>Employee ID</SortButton></th>
              <th><SortButton columnKey="name" setSort={setSort} sort={sort}>Name</SortButton></th>
              <th><SortButton columnKey="dept" setSort={setSort} sort={sort}>Dept</SortButton></th>
              <th><SortButton columnKey="position" setSort={setSort} sort={sort}>Position</SortButton></th>
              <th><SortButton columnKey="shift" setSort={setSort} sort={sort}>Shift</SortButton></th>
              <th><SortButton columnKey="shiftStart" setSort={setSort} sort={sort}>Shift Start</SortButton></th>
              <th><SortButton columnKey="scanIn" setSort={setSort} sort={sort}>Scan In</SortButton></th>
              <th><SortButton columnKey="status" setSort={setSort} sort={sort}>Status</SortButton></th>
              <th><SortButton columnKey="minutesLate" setSort={setSort} sort={sort}>Minutes Late</SortButton></th>
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
  deptFilter = "all",
  query = "",
  reportData,
  setDeptFilter,
  setQuery,
  setStatusFilter,
  standalone = false,
  statusFilter = "all",
}: {
  deptFilter?: string;
  query?: string;
  reportData: ReportData | null;
  setDeptFilter?: (value: string) => void;
  setQuery?: (value: string) => void;
  setStatusFilter?: (value: string) => void;
  standalone?: boolean;
  statusFilter?: string;
}) {
  const [page, setPage] = useState(1);
  const [sort, setSort_] = useState<SortState>(null);
  const setSort = setSort_ as (sort: SortState) => void;
  const sourceRows = reportData?.records.filter((record) => record.status !== "Absent") ?? [];
  const deptOptions = Array.from(new Set(sourceRows.map((record) => record.dept))).sort();
  const normalizedQuery = query.trim().toLowerCase();
  const allRows = sourceRows.filter((record) => {
    const matchesQuery = !normalizedQuery || [
      record.empId,
      record.name,
      record.dept,
      record.position,
      record.scanIn,
      record.status,
    ].some((value) => String(value).toLowerCase().includes(normalizedQuery));
    const matchesDept = deptFilter === "all" || record.dept === deptFilter;
    const matchesStatus = statusFilter === "all" || record.status === statusFilter;
    return matchesQuery && matchesDept && matchesStatus;
  });
  const sortedRows = sortAttendanceRows(allRows, sort, reportData?.monthlyLateCounts ?? {});
  const totalPages = Math.max(1, Math.ceil(sortedRows.length / pageSize));
  const safePage = Math.min(page, totalPages);
  const rows = sortedRows.slice((safePage - 1) * pageSize, safePage * pageSize);

  function updateFilter(callback: () => void) {
    callback();
    setPage(1);
  }

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
      {standalone ? (
        <div className="table-filters">
          <input
            aria-label="ค้นหา"
            placeholder="ค้นหา รหัส ชื่อ หน่วยงาน สถานะ"
            type="search"
            value={query}
            onChange={(event) => updateFilter(() => setQuery?.(event.target.value))}
          />
          <select
            aria-label="หน่วยงาน"
            value={deptFilter}
            onChange={(event) => updateFilter(() => setDeptFilter?.(event.target.value))}
          >
            <option value="all">ทุกหน่วยงาน</option>
            {deptOptions.map((dept) => (
              <option key={dept} value={dept}>{dept}</option>
            ))}
          </select>
          <select
            aria-label="สถานะ"
            value={statusFilter}
            onChange={(event) => updateFilter(() => setStatusFilter?.(event.target.value))}
          >
            <option value="all">ทุกสถานะ</option>
            <option value="Present">Present</option>
            <option value="Late">Late</option>
          </select>
          <button
            className="ghost-button"
            onClick={() => updateFilter(() => {
              setQuery?.("");
              setDeptFilter?.("all");
              setStatusFilter?.("all");
            })}
            type="button"
          >
            Clear
          </button>
        </div>
      ) : null}
      <div className="table-scroll">
        <table className="table data-table">
          <thead>
            <tr>
              <th>No.</th>
              <th><SortButton columnKey="empId" setSort={setSort} sort={sort}>รหัสพนักงาน</SortButton></th>
              <th><SortButton columnKey="name" setSort={setSort} sort={sort}>ชื่อ-สกุล</SortButton></th>
              <th><SortButton columnKey="dept" setSort={setSort} sort={sort}>หน่วยงาน</SortButton></th>
              <th><SortButton columnKey="position" setSort={setSort} sort={sort}>ตำแหน่ง</SortButton></th>
              <th>สถานีงานที่จัดสรร</th>
              <th>ระดับ Skill</th>
              <th><SortButton columnKey="scanIn" setSort={setSort} sort={sort}>เวลาเข้า</SortButton></th>
              <th><SortButton columnKey="status" setSort={setSort} sort={sort}>สถานะ</SortButton></th>
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

  const windowSize = 5;
  const halfWindow = Math.floor(windowSize / 2);
  let winStart = Math.max(1, page - halfWindow);
  const winEnd = Math.min(totalPages, winStart + windowSize - 1);
  if (winEnd - winStart + 1 < windowSize) {
    winStart = Math.max(1, winEnd - windowSize + 1);
  }
  const pages = Array.from({ length: winEnd - winStart + 1 }, (_, i) => winStart + i);

  return (
    <div className="pagination">
      <span>{start}-{end} จาก {totalRows} รายการ</span>
      <button disabled={!setPage || page <= 1} onClick={() => setPage?.(page - 1)} type="button">‹</button>
      {winStart > 1 ? <span>…</span> : null}
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
      {winEnd < totalPages ? <span>…</span> : null}
      <button disabled={!setPage || page >= totalPages} onClick={() => setPage?.(page + 1)} type="button">›</button>
    </div>
  );
}

function ReportDashboard({
  deptFilter,
  isLoadingReport,
  loadReportDashboard,
  query,
  reportData,
  selectedDept,
  setDeptFilter,
  setQuery,
  setSelectedDept,
  warnCountMap,
}: {
  deptFilter: string;
  isLoadingReport: boolean;
  loadReportDashboard: () => Promise<void>;
  query: string;
  reportData: ReportData | null;
  selectedDept: string;
  setDeptFilter: (value: string) => void;
  setQuery: (value: string) => void;
  setSelectedDept: (value: string) => void;
  warnCountMap: Record<string, number>;
}) {
  const [sort, setSort_] = useState<SortState>(null);
  const setSort = setSort_ as (sort: SortState) => void;
  const [tableStatusFilter, setTableStatusFilter] = useState<"all" | "Late" | "Absent">("all");

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
    monthlyLateCounts: {},
    isoTargetDate: "",
    targetMonthKey: "",
  };

  const scopedRecords = selectedDept === "all"
    ? data.records
    : data.records.filter((row) => row.dept === selectedDept);
  const scopedTotal = scopedRecords.length;
  const scopedPresent = scopedRecords.filter((row) => row.status === "Present").length;
  const scopedLate = scopedRecords.filter((row) => row.status === "Late").length;
  const scopedAbsent = scopedRecords.filter((row) => row.status === "Absent").length;
  const scopedCameToWork = scopedPresent + scopedLate;
  const lateRate = scopedCameToWork
    ? ((scopedLate / scopedCameToWork) * 100).toFixed(1)
    : "0.0";
  const maxDeptTotal = Math.max(...data.deptRows.map((row) => row.total), 1);
  const presentPercent = scopedTotal ? (scopedPresent / scopedTotal) * 100 : 0;
  const latePercent = scopedTotal ? (scopedLate / scopedTotal) * 100 : 0;
  const absentPercent = scopedTotal ? (scopedAbsent / scopedTotal) * 100 : 0;
  const lateDeptRows = Array.from(
    data.lateRows.reduce((map, row) => {
      map.set(row.dept, (map.get(row.dept) ?? 0) + 1);
      return map;
    }, new Map<string, number>()),
  )
    .map(([dept, count]) => ({ dept, count }))
    .sort((a, b) => b.count - a.count);
  const lateDeptTotal = Math.max(lateDeptRows.reduce((sum, row) => sum + row.count, 0), 1);
  const pieColors = ["#2563eb", "#0f172a", "#10b981", "#f59e0b", "#dc2626", "#7c3aed"];
  const selectedDeptLabel = selectedDept === "all" ? "ทั้งโรงงาน" : selectedDept;
  const normalizedQuery = query.trim().toLowerCase();
  const tableSourceRows = scopedRecords.filter((r) => r.status === "Late" || r.status === "Absent");
  const tableDeptOptions = Array.from(new Set(tableSourceRows.map((r) => r.dept))).sort();
  const filteredTableRows = tableSourceRows.filter((row) => {
    const matchesQuery = !normalizedQuery || [
      row.empId, row.name, row.dept, row.position, row.scanIn, row.status,
    ].some((v) => String(v).toLowerCase().includes(normalizedQuery));
    const matchesDept = deptFilter === "all" || row.dept === deptFilter;
    const matchesStatus = tableStatusFilter === "all" || row.status === tableStatusFilter;
    return matchesQuery && matchesDept && matchesStatus;
  });
  const sortedTableRows = sortAttendanceRows(filteredTableRows, sort, data.monthlyLateCounts);

  return (
    <section className="report-page">
      <div className="report-topstrip">
        <div className="report-topstrip-left">
          <CalendarDays size={16} />
          <span>ข้อมูลวันที่ <strong>{data.targetDate}</strong></span>
          {data.deptRows.length > 0 ? (
            <select
              className="dept-filter-select"
              value={selectedDept}
              onChange={(e) => { setSelectedDept(e.target.value); setDeptFilter("all"); }}
            >
              <option value="all">ทุกหน่วยงาน</option>
              {data.deptRows.map((r) => (
                <option key={r.dept} value={r.dept}>{r.dept}</option>
              ))}
            </select>
          ) : null}
        </div>
        <button
          className="primary-button report-refresh"
          disabled={isLoadingReport}
          onClick={loadReportDashboard}
          type="button"
        >
          <BarChart3 size={16} />
          {isLoadingReport ? "กำลังโหลด..." : "โหลดข้อมูล"}
        </button>
      </div>

      <section className="kpi-grid">
        <KpiCard
          icon={<UsersRound size={34} />}
          tone="green"
          label="พนักงานทั้งหมด"
          value={scopedTotal.toLocaleString()}
          unit="คน"
          note={selectedDeptLabel}
          progress={scopedTotal ? Math.round((scopedCameToWork / scopedTotal) * 100) : 0}
        />
        <KpiCard
          icon={<CheckCircle2 size={34} />}
          tone="blue"
          label="มาทำงาน"
          value={scopedPresent.toLocaleString()}
          unit="คน"
          note={`${presentPercent.toFixed(1)}% ของทั้งหมด`}
        />
        <KpiCard
          icon={<Clock size={34} />}
          tone="amber"
          label="มาสาย"
          value={scopedLate.toLocaleString()}
          unit="คน"
          note={`อัตราสาย ${lateRate}%`}
        />
        <KpiCard
          icon={<UserX size={34} />}
          tone="purple"
          label="ขาดงาน"
          value={scopedAbsent.toLocaleString()}
          unit="คน"
          note={`${absentPercent.toFixed(1)}% ของทั้งหมด`}
        />
        <DonutKpiCard
          present={scopedPresent}
          late={scopedLate}
          absent={scopedAbsent}
          total={scopedTotal}
          totalActive={scopedCameToWork}
        />
      </section>

      <section className="report-grid">
        <div className="panel report-card">
          <div className="panel-title-row">
            <h3>การเข้างานรายแผนก</h3>
            {selectedDept !== "all" ? (
              <button className="ghost-button" onClick={() => setSelectedDept("all")} type="button">
                ดูทั้งหมด
              </button>
            ) : null}
          </div>
          <div className="stack-legend">
            <span><i className="present" />Present</span>
            <span><i className="late" />Late</span>
            <span><i className="absent" />Absent</span>
          </div>
          <div className="stacked-bars">
            {data.deptRows.map((row) => (
              <button
                className={`stacked-row dept-click ${selectedDept === row.dept ? "active" : ""}`}
                key={row.dept}
                onClick={() => {
                  setSelectedDept(row.dept);
                  setDeptFilter("all");
                }}
                type="button"
              >
                <span className="stacked-dept-name">{row.dept}</span>
                <div className="stacked-track">
                  <i className="present" style={{ width: `${(row.present / maxDeptTotal) * 100}%` }} />
                  <i className="late" style={{ width: `${(row.late / maxDeptTotal) * 100}%` }} />
                  <i className="absent" style={{ width: `${(row.absent / maxDeptTotal) * 100}%` }} />
                </div>
                <div className="stacked-row-end">
                  <strong>{row.total}</strong>
                  <span className="stacked-mini-badges">
                    {row.late > 0 ? <span className="mini-badge late">{row.late}L</span> : null}
                    {row.absent > 0 ? <span className="mini-badge absent">{row.absent}A</span> : null}
                  </span>
                </div>
              </button>
            ))}
            {data.deptRows.length === 0 ? <p className="empty-copy">ยังไม่มีข้อมูล report</p> : null}
          </div>
        </div>

        <div className="panel report-overview-card">
          <h3>ภาพรวมการเข้างาน{selectedDept !== "all" ? ` · ${selectedDept}` : ""}</h3>
          <div className="overview-donut-row">
            <div
              className="report-donut"
              style={{
                background: `conic-gradient(#10b981 0 ${presentPercent}%, #f59e0b ${presentPercent}% ${presentPercent + latePercent}%, #dc2626 ${presentPercent + latePercent}% 100%)`,
              }}
            >
              <div>
                <strong>{scopedTotal}</strong>
                <span>คน</span>
              </div>
            </div>
            <div className="report-legend">
              <LegendRow color="green" label="Present" value={String(scopedPresent)} percent={`${presentPercent.toFixed(1)}%`} />
              <LegendRow color="amber" label="Late" value={String(scopedLate)} percent={`${latePercent.toFixed(1)}%`} />
              <LegendRow color="red" label="Absent" value={String(scopedAbsent)} percent={`${absentPercent.toFixed(1)}%`} />
            </div>
          </div>
          <div className="late-dept-breakdown">
            <h4>หน่วยงานที่มาสาย</h4>
            {lateDeptRows.slice(0, 6).map((row, index) => (
              <div className="late-dept-row" key={row.dept}>
                <i style={{ background: pieColors[index % pieColors.length] }} />
                <span className="late-dept-name">{row.dept}</span>
                <div className="late-dept-bar">
                  <div style={{ width: `${(row.count / lateDeptTotal) * 100}%`, background: pieColors[index % pieColors.length] }} />
                </div>
                <span className="late-dept-count">{row.count}</span>
              </div>
            ))}
            {lateDeptRows.length === 0 ? <p className="empty-copy">ยังไม่มีข้อมูลคนมาสาย</p> : null}
          </div>
        </div>
      </section>

      <section className="panel report-table-panel">
        <div className="report-table-header">
          <div className="report-table-header-top">
            <div className="report-table-title-row">
              <h3>รายละเอียด Late &amp; Absent</h3>
              {selectedDept !== "all" && <span className="dept-filter-chip">{selectedDept}</span>}
              <span className="table-count-badge">{sortedTableRows.length} คน</span>
            </div>
            <div className="report-table-actions">
              <div className="status-tabs">
                {(["all", "Late", "Absent"] as const).map((s) => (
                  <button
                    key={s}
                    className={`status-tab${tableStatusFilter === s ? " active" : ""}${s === "Late" ? " amber" : s === "Absent" ? " red" : ""}`}
                    onClick={() => setTableStatusFilter(s)}
                    type="button"
                  >
                    {s === "all" ? `ทั้งหมด` : s}
                    <span className="tab-count">{s === "all" ? tableSourceRows.length : s === "Late" ? scopedLate : scopedAbsent}</span>
                  </button>
                ))}
              </div>
              <button
                className="primary-button small"
                disabled={filteredTableRows.length === 0}
                onClick={() => exportLateAbsentToExcel(filteredTableRows, data.monthlyLateCounts, selectedDeptLabel)}
                type="button"
              >
                <Download size={13} />
                Export
              </button>
            </div>
          </div>
          <div className="report-table-filters-row">
            <input
              aria-label="ค้นหา"
              placeholder="ค้นหา ชื่อ / รหัสพนักงาน / หน่วยงาน"
              type="search"
              value={query}
              onChange={(event) => setQuery(event.target.value)}
            />
            <select
              aria-label="หน่วยงาน"
              value={deptFilter}
              onChange={(event) => setDeptFilter(event.target.value)}
            >
              <option value="all">ทุกหน่วยงาน</option>
              {tableDeptOptions.map((dept) => (
                <option key={dept} value={dept}>{dept}</option>
              ))}
            </select>
            {(query || deptFilter !== "all") && (
              <button
                className="ghost-button"
                onClick={() => { setQuery(""); setDeptFilter("all"); }}
                type="button"
              >
                <X size={13} /> ล้าง
              </button>
            )}
          </div>
        </div>
        <div className="table-scroll">
          <table className="table">
            <thead>
              <tr>
                <th>No.</th>
                <th><SortButton columnKey="dept" setSort={setSort} sort={sort}>หน่วยงาน</SortButton></th>
                <th><SortButton columnKey="name" setSort={setSort} sort={sort}>ชื่อ-สกุล</SortButton></th>
                <th><SortButton columnKey="shift" setSort={setSort} sort={sort}>กะ</SortButton></th>
                <th><SortButton columnKey="shiftStart" setSort={setSort} sort={sort}>เวลาเข้างาน</SortButton></th>
                <th><SortButton columnKey="scanIn" setSort={setSort} sort={sort}>Scan In</SortButton></th>
                <th><SortButton columnKey="status" setSort={setSort} sort={sort}>สถานะ</SortButton></th>
                <th><SortButton columnKey="minutesLate" setSort={setSort} sort={sort}>สาย</SortButton></th>
                <th><SortButton columnKey="monthlyLate" setSort={setSort} sort={sort}>สายเดือนนี้</SortButton></th>
                <th>เสี่ยง</th>
                <th>เตือนแล้ว</th>
              </tr>
            </thead>
            <tbody key={tableStatusFilter}>
              {sortedTableRows.map((row, index) => {
                const monthly = data.monthlyLateCounts[row.empId] ?? 0;
                const isRisk = monthly >= 3;
                const warnCount = warnCountMap[row.empId] ?? 0;
                return (
                  <tr key={`${row.empId}-${row.scanIn}-rpt`} className={isRisk ? "row-risk" : ""}>
                    <td>{index + 1}</td>
                    <td><span className="dept-chip">{row.dept}</span></td>
                    <td>{row.name}</td>
                    <td>{row.shift}</td>
                    <td>{row.shiftStart}</td>
                    <td>{row.scanIn ?? "-"}</td>
                    <td><span className={`status-pill ${row.status.toLowerCase()}`}>{row.status}</span></td>
                    <td>{row.status !== "Absent" ? <span className="late-minutes-badge">{formatLateTime(row.minutesLate)}</span> : "-"}</td>
                    <td>
                      <span className={monthly >= 3 ? "monthly-late-high" : ""}>
                        {monthly > 0 ? monthly : "-"}
                      </span>
                    </td>
                    <td>{isRisk ? <span className="risk-badge">เสี่ยง</span> : null}</td>
                    <td>
                      {warnCount > 0
                        ? <span className="warn-count-badge">✓ {warnCount} ครั้ง</span>
                        : <span className="no-warn">-</span>}
                    </td>
                  </tr>
                );
              })}
              {sortedTableRows.length === 0 ? (
                <tr><td colSpan={11}>{reportData ? "ไม่มีข้อมูลตามเงื่อนไขที่เลือก" : "กด โหลดข้อมูล เพื่อสร้างรายงาน"}</td></tr>
              ) : null}
            </tbody>
          </table>
        </div>
      </section>
    </section>
  );
}

function ReportMetric({
  label,
  tone,
  value,
  sublabel,
  isRate,
  icon,
}: {
  label: string;
  tone?: "green" | "amber" | "red" | "purple";
  value: number | string;
  sublabel?: string;
  isRate?: boolean;
  icon?: ReactNode;
}) {
  return (
    <div className={`report-metric${tone ? ` ${tone}` : ""}${isRate ? " is-rate" : ""}`}>
      {icon && <div className="report-metric-icon">{icon}</div>}
      <strong className="report-metric-value">{value}</strong>
      <span className="report-metric-label">{label}</span>
      {sublabel && <span className="report-metric-sub">{sublabel}</span>}
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
    <article className={`kpi-card kpi-${tone}`}>
      <div className={`kpi-icon ${tone}`}>{icon}</div>
      <div className="kpi-body">
        <span>{label}</span>
        <div>
          <strong>{value}</strong>
          <b>{unit}</b>
        </div>
        <p>{note}</p>
        {progress ? (
          <div className="progress-wrap">
            <div className="progress-line">
              <i style={{ width: `${progress}%` }} />
            </div>
            <em>{progress}%</em>
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
