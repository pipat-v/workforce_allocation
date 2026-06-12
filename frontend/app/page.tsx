"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import {
  AlertTriangle,
  BarChart3,
  BriefcaseBusiness,
  CalendarDays,
  CalendarOff,
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
  dept: string;
  dayoff: string;
  shift: string;
  shiftStart: string;
  raw: Record<string, unknown>;
};

type SkillMatrixSaveRow = {
  empId: string;
  skill: string;
  level: number;
};

type HolidayRow = {
  id: string;
  date: string;
  name: string;
  type: "buddhist_holy_day" | "public_holiday" | "company_holiday";
};

type SkillFlatRow = {
  id: string;
  empId: string;
  name: string;
  dept: string;
  skill: string;
  level: number;
  origLevel: number;
};

type AttendanceRecord = {
  empId: string;
  name: string;
  dept: string;
  position: string;
  shift: string;
  shiftStart: string;
  scanIn: string;
  status: "Present" | "Late" | "Absent" | "DayOff";
  minutesLate: number;
};

type DailyStat = {
  isoDate: string;
  present: number;
  late: number;
  absent: number;
  total: number;
};

type ReportData = {
  targetDate: string;
  isoTargetDate: string;
  targetMonthKey: string;
  totalEmployees: number;
  present: number;
  late: number;
  absent: number;
  dayoff: number;
  deptRows: Array<{ dept: string; present: number; late: number; absent: number; dayoff: number; total: number }>;
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

type MasterFileKey = (typeof masterFileTypes)[number]["key"];
type MasterUploadState = Record<MasterFileKey, File | null>;
type SortDirection = "asc" | "desc";
type SortState = { key: string; direction: SortDirection } | null;

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
  const [warnDates, setWarnDates] = useState<Record<string, string[]>>({});
  const [monthlyLateMinutes, setMonthlyLateMinutes] = useState<Record<string, number>>({});
  const [dailyStats, setDailyStats] = useState<DailyStat[]>([]);
  const [prevMonthLateCounts, setPrevMonthLateCounts] = useState<Record<string, number>>({});
  const [selectedRunId, setSelectedRunId] = useState<string | null>(null);
  const [showRunPicker, setShowRunPicker] = useState(false);
  const [holidayDates, setHolidayDates] = useState<Set<string>>(() => {
    const all = new Set<string>();
    for (const s of Object.values(buddhistHolyDaysByYear)) for (const d of s) all.add(d);
    return all;
  });

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

  const latestRun = (selectedRunId ? runs.find(r => r.id === selectedRunId) : null) ?? runs[0];
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
      dayoff: filtered.filter((r) => r.status === "DayOff").length,
      lateRows: [...lateFiltered].sort((a, b) => b.minutesLate - a.minutesLate),
      deptRows: reportData.deptRows.filter((r) => r.dept === dashboardDeptFilter),
    };
  }, [reportData, dashboardDeptFilter]);

  const totalEmployees = dashboardReport?.totalEmployees ?? 0;
  const presentPeople = dashboardReport?.present ?? 0;
  const latePeople = dashboardReport?.late ?? 0;
  const absentPeople = dashboardReport?.absent ?? 0;
  const dayoffPeople = dashboardReport?.dayoff ?? 0;
  const totalActivePeople = presentPeople + latePeople;
  const presentRate = totalEmployees ? Math.round((totalActivePeople / totalEmployees) * 100) : 0;
  const pctPresent = totalEmployees ? Math.round((presentPeople / totalEmployees) * 100) : 0;
  const pctLate = totalEmployees ? Math.round((latePeople / totalEmployees) * 100) : 0;
  const pctAbsent = totalEmployees ? Math.round((absentPeople / totalEmployees) * 100) : 0;
  const pctDayoff = totalEmployees ? Math.round((dayoffPeople / totalEmployees) * 100) : 0;
  const todayDate = new Date().toLocaleDateString("th-TH", { day: "numeric", month: "long", year: "numeric" });
  const workDateBase = latestRun?.target_date ? new Date(latestRun.target_date) : new Date();
  const workDate = workDateBase.toLocaleDateString("th-TH", {
    day: "numeric",
    month: "long",
    year: "numeric",
  });
  const [workTime, setWorkTime] = useState(() =>
    new Date().toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit" }),
  );
  const datePickerRef = useRef<HTMLDivElement>(null);
  const selectedDateKey = latestRun?.target_date ?? latestRun?.created_at?.slice(0, 10) ?? "";
  const availableRunsByDate = useMemo(() => {
    const map = new Map<string, AllocationRun>();
    for (const run of runs) {
      if (!run.scan_file_path) continue;
      const key = run.target_date ?? run.created_at.slice(0, 10);
      if (!map.has(key)) map.set(key, run);
    }
    return map;
  }, [runs]);

  useEffect(() => {
    if (!showRunPicker) return;
    function handleClickOutside(e: MouseEvent) {
      if (datePickerRef.current && !datePickerRef.current.contains(e.target as Node)) {
        setShowRunPicker(false);
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, [showRunPicker]);

  useEffect(() => {
    const tick = () =>
      setWorkTime(new Date().toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit" }));
    const id = setInterval(tick, 30_000);
    return () => clearInterval(id);
  }, []);

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

  useEffect(() => {
    supabase
      .from("holidays")
      .select("date")
      .then(({ data }) => {
        if (data && data.length > 0) {
          setHolidayDates(new Set(data.map((r: { date: string }) => r.date)));
        }
      });
  }, []);

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
      .limit(100);

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
      "เวลาเข้างาน": row.shiftStart,
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

  async function saveSkillMatrixRows(rows: SkillMatrixSaveRow[]) {
    setError("");
    setMessage("");

    const fileId = crypto.randomUUID();
    const path = `${publicWorkspace}/masters/skill_matrix/${fileId}.xlsx`;
    const workbookRows = rows.map((row) => ({
      "Employee ID": row.empId,
      "Skill": row.skill,
      "Level": row.level,
      "Can Do": row.level > 0 ? 1 : 0,
    }));
    const worksheet = XLSX.utils.json_to_sheet(workbookRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "SkillMatrix");
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
      .eq("file_type", "skill_matrix");

    if (deactivateError) {
      setError(deactivateError.message);
      return;
    }

    const { error: insertError } = await supabase
      .from("master_data_files")
      .insert({
        owner_id: null,
        file_type: "skill_matrix",
        file_path: path,
        original_filename: "SkillMatrix-edited.xlsx",
        is_active: true,
      });

    if (insertError) {
      setError(insertError.message);
      return;
    }

    setMessage("บันทึก Skill Matrix master แล้ว");
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
    if (run.scan_file_path) {
      const { error: storageError } = await supabase.storage
        .from("workforce-inputs")
        .remove([run.scan_file_path]);
      if (storageError) {
        setError(storageError.message);
        return;
      }
    }
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
      const latestRun = (selectedRunId ? runs.find(r => r.id === selectedRunId) : null) ?? runs.find((run) => run.scan_file_path);

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

      const latestReport = buildReportData(employeeRows, scanRows, manpowerRows, dayoffShiftRows, holidayDates);
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
      const lateMinutesAcc: Record<string, number> = {};
      const prevMonthLateAcc: Record<string, number> = {};
      const dailyStatsAcc: Record<string, DailyStat> = {};
      const [tY, tM] = latestReport.targetMonthKey.split("-").map(Number);
      const prevMonthKey = tM === 1
        ? `${tY - 1}-12`
        : `${tY}-${String(tM - 1).padStart(2, "0")}`;

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
        const dayReport = buildReportData(employeeRows, rows, manpowerRows, dayoffShiftRows, holidayDates);

        if (dayReport.targetMonthKey === latestReport.targetMonthKey) {
          for (const lateRow of dayReport.lateRows) {
            monthlyLateCounts[lateRow.empId] = (monthlyLateCounts[lateRow.empId] ?? 0) + 1;
            lateMinutesAcc[lateRow.empId] = (lateMinutesAcc[lateRow.empId] ?? 0) + lateRow.minutesLate;
          }
          if (dayReport.isoTargetDate && !dailyStatsAcc[dayReport.isoTargetDate]) {
            dailyStatsAcc[dayReport.isoTargetDate] = {
              isoDate: dayReport.isoTargetDate,
              present: dayReport.present,
              late: dayReport.late,
              absent: dayReport.absent,
              total: dayReport.totalEmployees,
            };
          }
        } else if (dayReport.targetMonthKey === prevMonthKey) {
          for (const lateRow of dayReport.lateRows) {
            prevMonthLateAcc[lateRow.empId] = (prevMonthLateAcc[lateRow.empId] ?? 0) + 1;
          }
        }
      }

      setMonthlyLateMinutes(lateMinutesAcc);
      setPrevMonthLateCounts(prevMonthLateAcc);
      setDailyStats(Object.values(dailyStatsAcc).sort((a, b) => a.isoDate.localeCompare(b.isoDate)));
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
    const dates: Record<string, string[]> = {};
    for (const r of rows as Array<{ emp_id: string; warn_date: string }>) {
      map[r.emp_id] = (map[r.emp_id] ?? 0) + 1;
      dates[r.emp_id] = [...(dates[r.emp_id] ?? []), r.warn_date].sort();
    }
    setWarnCountMap(map);
    setWarnDates(dates);
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
            <div className="topbar-date-group">
              <div className="today-chip">
                <Clock size={14} />
                <div>
                  <span className="today-chip-label">วันนี้</span>
                  <strong>{todayDate}</strong>
                  <span className="work-time">{workTime} น.</span>
                </div>
              </div>
              <div className="date-picker-wrap" ref={datePickerRef}>
              <div className="date-picker" style={{ cursor: "pointer" }} onClick={() => setShowRunPicker(v => !v)}>
                <CalendarDays size={19} />
                <div>
                  <div className="date-picker-label-row">
                    <span>ข้อมูลวันที่</span>
                  </div>
                  <strong>{workDate}</strong>
                </div>
                <ChevronDown size={17} style={{ transition: "transform 0.2s", transform: showRunPicker ? "rotate(180deg)" : "rotate(0deg)" }} />
              </div>
              {showRunPicker && (
                <div className="run-picker-dropdown cal-picker-dropdown">
                  <CalendarPicker
                    value={selectedDateKey}
                    availableDates={availableRunsByDate}
                    onChange={(dateStr) => {
                      const run = availableRunsByDate.get(dateStr);
                      if (run) { setSelectedRunId(run.id); setLoadedReportKey(""); setShowRunPicker(false); }
                    }}
                  />
                </div>
              )}
            </div>
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
            {activeTab === "report" && allDeptOptions.length > 0 ? (
              <select
                aria-label="กรองหน่วยงาน"
                className="dept-filter-select"
                value={selectedReportDept}
                onChange={(e) => setSelectedReportDept(e.target.value)}
              >
                <option value="all">ทุกหน่วยงาน</option>
                {allDeptOptions.map((dept) => (
                  <option key={dept} value={dept}>{dept}</option>
                ))}
              </select>
            ) : null}
            {activeTab === "report" ? (
              <button
                className="primary-button"
                style={{ height: 36, fontSize: 13 }}
                disabled={isLoadingReport}
                onClick={loadReportDashboard}
                type="button"
              >
                <BarChart3 size={15} />
                {isLoadingReport ? "กำลังโหลด..." : "โหลดข้อมูล"}
              </button>
            ) : null}
          </div>
          {(message || error) ? (
            <div className={`toast ${error ? "error" : ""}`}>
              <span>{error || message}</span>
              <button
                className="toast-close"
                onClick={() => { setError(""); setMessage(""); }}
                type="button"
                aria-label="ปิด"
              >
                <X size={14} />
              </button>
            </div>
          ) : null}
        </section>

        {activeTab === "dashboard" ? (
          <>
            <section className="kpi-grid kpi-grid-5">
              <KpiCard
                icon={<UsersRound size={34} />}
                tone="green"
                label="มาทำงาน"
                value={totalActivePeople.toLocaleString()}
                unit="คน"
                note={`จากทั้งหมด ${totalEmployees.toLocaleString()} คน`}
                progress={presentRate}
              />
              <KpiCard
                icon={<ClipboardCheck size={34} />}
                tone="blue"
                label="ตรงเวลา"
                value={presentPeople.toLocaleString()}
                unit="คน"
                note={`${pctPresent}% ของพนักงานทั้งหมด`}
              />
              <KpiCard
                icon={<Clock size={34} />}
                tone="amber"
                label="มาสาย"
                value={latePeople.toLocaleString()}
                unit="คน"
                note={`${pctLate}% ของพนักงานทั้งหมด`}
              />
              <KpiCard
                icon={<UserX size={34} />}
                tone="purple"
                label="ขาดงาน"
                value={absentPeople.toLocaleString()}
                unit="คน"
                note={`${pctAbsent}% ของพนักงานทั้งหมด`}
              />
              <KpiCard
                icon={<CalendarOff size={34} />}
                tone="gray"
                label="วันหยุด"
                value={dayoffPeople.toLocaleString()}
                unit="คน"
                note={`${pctDayoff}% ของพนักงานทั้งหมด`}
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
              warnCountMap={warnCountMap}
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
            onHolidaysChanged={(dates) => setHolidayDates(dates)}
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
            dailyStats={dailyStats}
            deptFilter={reportLateDept}
            monthlyLateMinutes={monthlyLateMinutes}
            prevMonthLateCounts={prevMonthLateCounts}
            query={reportLateQuery}
            reportData={reportData}
            selectedDept={selectedReportDept}
            setDeptFilter={setReportLateDept}
            setQuery={setReportLateQuery}
            setSelectedDept={setSelectedReportDept}
            warnCountMap={warnCountMap}
            warnDates={warnDates}
          />
        ) : null}

        {activeTab === "skill" ? (
          <SkillMatrixPage
            activeFile={activeMasterMap.skill_matrix}
            employeeMasterFile={activeMasterMap.employee_master}
            saveSkillMatrixRows={saveSkillMatrixRows}
          />
        ) : null}

        {!["dashboard", "master", "timestamp", "results", "timestamp_dept", "report", "skill"].includes(activeTab) ? (
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
  holidaySet?: Set<string>,
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

  const records: AttendanceRecord[] = baseEmployees.flatMap((employee): AttendanceRecord[] => {
    const dayoffShift = dayoffShiftMap.get(employee.empId);
    const scans = scanByEmp.get(employee.empId)?.times ?? [];
    const scanIn = scans.sort((a, b) => a.getTime() - b.getTime())[0];
    const shift = normalizeShiftLabel(dayoffShift?.shift) || "กะ 1";
    const shiftStart =
      deptShiftStart.get(makeDeptShiftKey(employee.dept, shift)) ??
      deptShiftStart.get(makeDeptShiftKey(employee.dept, "")) ??
      "07:00";
    const isScheduledOff = latestTimestamp
      ? isEmployeeDayOff(dayoffShift?.dayoff, latestTimestamp, holidaySet)
      : false;
    if (isScheduledOff) return [{
      empId: employee.empId,
      name: employee.name,
      dept: employee.dept,
      position: employee.position,
      shift,
      shiftStart,
      scanIn: "-",
      status: "DayOff" as const,
      minutesLate: 0,
    }];

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

  const deptMap = new Map<string, { dept: string; present: number; late: number; absent: number; dayoff: number; total: number }>();
  for (const record of records) {
    const current = deptMap.get(record.dept) ?? {
      dept: record.dept,
      present: 0,
      late: 0,
      absent: 0,
      dayoff: 0,
      total: 0,
    };
    current.total += 1;
    if (record.status === "Present") current.present += 1;
    if (record.status === "Late") current.late += 1;
    if (record.status === "Absent") current.absent += 1;
    if (record.status === "DayOff") current.dayoff += 1;
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
    dayoff: records.filter((record) => record.status === "DayOff").length,
    deptRows: Array.from(deptMap.values())
      .sort((a, b) => b.total - a.total),
    lateRows: records
      .filter((record) => record.status === "Late")
      .sort((a, b) => b.minutesLate - a.minutesLate),
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
  const shiftStartRaw = row["เวลาเข้างาน"] ?? row["เวลาเข้า"] ?? row["shift_start"];
  return {
    id: `${empId || "row"}-${index}`,
    empId,
    name: `${firstName} ${lastName}`.trim() || fallbackName || empId,
    dept: findRowCol(row, "หน่วยงาน", "Org. Unit Description", "Name (Section)", "แผนก", "Department"),
    dayoff: findRowCol(row, "วันหยุดประจำสัปดาห์", "วันหยุด", "dayoff", "Dayoff", "Day Off"),
    shift: findRowCol(row, "อยู่กะไหน", "shift", "กะ", "Shift"),
    shiftStart: normalizeTimeText(shiftStartRaw),
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

function addHoursToTime(timeStr: string, hours: number): string {
  if (!timeStr) return "";
  const [h, m] = timeStr.split(":").map(Number);
  const total = h * 60 + m + hours * 60;
  const nh = Math.floor(total / 60) % 24;
  const nm = total % 60;
  return `${String(nh).padStart(2, "0")}:${String(nm).padStart(2, "0")}`;
}

function isHolidayDate(date: Date, holidaySet?: Set<string>) {
  if (holidaySet) return holidaySet.has(toDateKey(date));
  return isBuddhistHolyDay(date);
}

function isEmployeeDayOff(dayoff: string | undefined, targetDate: Date, holidaySet?: Set<string>) {
  const value = String(dayoff ?? "").trim();
  if (!value) return false;
  const todayCode = getThaiWeekdayCode(targetDate);
  // split by comma / slash / space / pipe to support "ส,อา" "ส/อา" "ส อา" etc.
  const parts = value.split(/[,/|\s]+/).map((s) => s.trim()).filter(Boolean);
  return parts.some((part) => part === "พระ" ? isHolidayDate(targetDate, holidaySet) : part === todayCode);
}

function getThaiWeekdayCode(date: Date) {
  return ["อา", "จ", "อ", "พ", "พฤ", "ศ", "ส"][date.getDay()];
}

const buddhistHolyDaysByYear: Record<string, Set<string>> = {
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
  "2028": new Set([
    "2028-01-01", "2028-01-09", "2028-01-16", "2028-01-24", "2028-01-31",
    "2028-02-08", "2028-02-15", "2028-02-23",
    "2028-03-01", "2028-03-09", "2028-03-16", "2028-03-24", "2028-03-31",
    "2028-04-07", "2028-04-14", "2028-04-22", "2028-04-30",
    "2028-05-07", "2028-05-14", "2028-05-21", "2028-05-28",
    "2028-06-05", "2028-06-12", "2028-06-20", "2028-06-27",
    "2028-07-05", "2028-07-12", "2028-07-19", "2028-07-26",
    "2028-08-03", "2028-08-10", "2028-08-17", "2028-08-24",
    "2028-09-01", "2028-09-08", "2028-09-15", "2028-09-22", "2028-09-30",
    "2028-10-07", "2028-10-14", "2028-10-21", "2028-10-29",
    "2028-11-05", "2028-11-12", "2028-11-19", "2028-11-27",
    "2028-12-04", "2028-12-11", "2028-12-18", "2028-12-26",
  ]),
  "2029": new Set([
    "2029-01-02", "2029-01-10", "2029-01-17", "2029-01-25",
    "2029-02-01", "2029-02-08", "2029-02-15", "2029-02-24",
    "2029-03-03", "2029-03-10", "2029-03-18", "2029-03-25",
    "2029-04-01", "2029-04-09", "2029-04-16", "2029-04-24",
    "2029-05-01", "2029-05-08", "2029-05-16", "2029-05-23", "2029-05-30",
    "2029-06-06", "2029-06-13", "2029-06-21", "2029-06-28",
    "2029-07-05", "2029-07-13", "2029-07-20", "2029-07-28",
    "2029-08-03", "2029-08-11", "2029-08-18", "2029-08-26",
    "2029-09-02", "2029-09-09", "2029-09-17", "2029-09-24",
    "2029-10-01", "2029-10-08", "2029-10-16", "2029-10-23", "2029-10-31",
    "2029-11-07", "2029-11-15", "2029-11-22", "2029-11-30",
    "2029-12-06", "2029-12-14", "2029-12-21", "2029-12-29",
  ]),
  "2030": new Set([
    "2030-01-05", "2030-01-12", "2030-01-20", "2030-01-27",
    "2030-02-03", "2030-02-10", "2030-02-18", "2030-02-25",
    "2030-03-04", "2030-03-11", "2030-03-19", "2030-03-26",
    "2030-04-02", "2030-04-09", "2030-04-17", "2030-04-24",
    "2030-05-01", "2030-05-08", "2030-05-16", "2030-05-23", "2030-05-31",
    "2030-06-07", "2030-06-14", "2030-06-22", "2030-06-29",
    "2030-07-06", "2030-07-13", "2030-07-21", "2030-07-28",
    "2030-08-04", "2030-08-11", "2030-08-19", "2030-08-26",
    "2030-09-02", "2030-09-09", "2030-09-17", "2030-09-24",
    "2030-10-01", "2030-10-08", "2030-10-16", "2030-10-23", "2030-10-31",
    "2030-11-07", "2030-11-14", "2030-11-22", "2030-11-29",
    "2030-12-06", "2030-12-13", "2030-12-21", "2030-12-28",
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
  warnCountMap: Record<string, number> = {},
) {
  if (key === "minutesLate") return row.minutesLate;
  if (key === "monthlyLate") return monthlyLateCounts[row.empId] ?? 0;
  if (key === "warnCount") return warnCountMap[row.empId] ?? 0;
  return (row as Record<string, unknown>)[key] as string ?? "";
}

function sortAttendanceRows(
  rows: AttendanceRecord[],
  sort: SortState,
  monthlyLateCounts: Record<string, number> = {},
  warnCountMap: Record<string, number> = {},
) {
  if (!sort) return rows;

  return [...rows].sort((a, b) => {
    const aValue = getAttendanceSortValue(a, sort.key, monthlyLateCounts, warnCountMap);
    const bValue = getAttendanceSortValue(b, sort.key, monthlyLateCounts, warnCountMap);

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

const STATUS_TH: Record<string, string> = {
  Present: "ตรงเวลา",
  Late: "มาสาย",
  Absent: "ขาดงาน",
  DayOff: "วันหยุด",
};

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
  dayoff = 0,
  total,
  totalActive,
}: {
  present: number;
  late: number;
  absent: number;
  dayoff?: number;
  total: number;
  totalActive: number;
}) {
  const presentPct = total ? (present / total) * 100 : 0;
  const latePct = total ? (late / total) * 100 : 0;
  const absentPct = total ? (absent / total) * 100 : 0;
  const dayoffPct = total ? (dayoff / total) * 100 : 0;

  return (
    <article className="kpi-card kpi-donut">
      <div className="kpi-donut-header">
        <span className="kpi-bar-label">{totalActive} คน</span>
        <div className="kpi-stacked-bar">
          <div className="kpi-bar-fill present" style={{ width: `${presentPct}%` }} />
          <div className="kpi-bar-fill late" style={{ width: `${latePct}%` }} />
          <div className="kpi-bar-fill absent" style={{ width: `${absentPct}%` }} />
          <div className="kpi-bar-fill dayoff" style={{ width: `${dayoffPct}%` }} />
        </div>
      </div>
      <div className="kpi-donut-legend">
        <LegendRow color="green" label="ตรงเวลา" value={String(present)} percent={`${presentPct.toFixed(1)}%`} />
        <LegendRow color="amber" label="มาสาย" value={String(late)} percent={`${latePct.toFixed(1)}%`} />
        <LegendRow color="red" label="ขาดงาน" value={String(absent)} percent={`${absentPct.toFixed(1)}%`} />
        <LegendRow color="gray" label="วันหยุด" value={String(dayoff)} percent={`${dayoffPct.toFixed(1)}%`} />
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
      "สถานะ": STATUS_TH[r.status] ?? r.status,
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
  warnCountMap,
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
  warnCountMap: Record<string, number>;
  warnedIds: Set<string>;
  warnPending: Set<string>;
}) {
  const [detailStatusFilter, setDetailStatusFilter] = useState("all");
  const [detailSort, setDetailSort_] = useState<SortState>(null);
  const setDetailSort = setDetailSort_ as (sort: SortState) => void;
  const [leaveMap, setLeaveMap] = useState<Map<string, string>>(new Map());
  const [warnPanelCollapsed, setWarnPanelCollapsed] = useState(true);
  const [detailPanelCollapsed, setDetailPanelCollapsed] = useState(true);

  useEffect(() => {
    if (!isoTargetDate) return;
    const absentIds = (reportData?.records ?? [])
      .filter((r: { status: string }) => r.status === "Absent")
      .map((r: { empId: string }) => r.empId);
    supabase.from("leave_records").select("emp_id, leave_type").eq("leave_date", isoTargetDate)
      .then(async ({ data: rows }) => {
        const map = new Map((rows ?? []).map((r: { emp_id: string; leave_type: string }) => [r.emp_id, r.leave_type]));
        const missing = absentIds.filter((id: string) => !map.has(id));
        if (missing.length) {
          await supabase.from("leave_records").upsert(
            missing.map((empId: string) => ({ emp_id: empId, leave_date: isoTargetDate, leave_type: "ขาดงาน" })),
            { onConflict: "emp_id,leave_date" }
          );
          missing.forEach((id: string) => map.set(id, "ขาดงาน"));
        }
        setLeaveMap(map);
      });
  }, [isoTargetDate, reportData]);

  const saveLeave = async (empId: string, leaveType: string) => {
    if (!isoTargetDate || !leaveType) return;
    setLeaveMap(prev => new Map(prev).set(empId, leaveType));
    await supabase.from("leave_records").upsert(
      { emp_id: empId, leave_date: isoTargetDate, leave_type: leaveType },
      { onConflict: "emp_id,leave_date" }
    );
  };

  const [confirmation, setConfirmation] = useState<{ confirmed_by: string; confirmed_at: string } | null | undefined>(undefined);
  const [confirmName, setConfirmName] = useState("");
  const [isConfirming, setIsConfirming] = useState(false);

  useEffect(() => {
    if (!isoTargetDate) { setConfirmation(undefined); return; }
    const deptKey = dashboardDeptFilter === "all" ? "ทุกหน่วยงาน" : dashboardDeptFilter;
    supabase.from("daily_confirmations")
      .select("confirmed_by, confirmed_at")
      .eq("confirm_date", isoTargetDate)
      .eq("dept", deptKey)
      .maybeSingle()
      .then(({ data }) => setConfirmation(data ?? null));
  }, [isoTargetDate, dashboardDeptFilter]);

  const handleConfirm = async () => {
    if (!confirmName.trim()) {
      const input = document.querySelector<HTMLInputElement>(".confirm-name-input");
      input?.focus();
      input?.classList.add("confirm-input-error");
      setTimeout(() => input?.classList.remove("confirm-input-error"), 1200);
      return;
    }
    if (!isoTargetDate) return;
    setIsConfirming(true);
    const deptKey = dashboardDeptFilter === "all" ? "ทุกหน่วยงาน" : dashboardDeptFilter;
    const leaveCounts: Record<string, number> = {};
    for (const lt of leaveMap.values()) leaveCounts[lt] = (leaveCounts[lt] ?? 0) + 1;
    const { data, error } = await supabase.from("daily_confirmations").upsert({
      confirm_date: isoTargetDate,
      dept: deptKey,
      confirmed_by: confirmName.trim(),
      late_count: reportData?.late ?? 0,
      absent_count: reportData?.absent ?? 0,
      leave_breakdown: leaveCounts,
    }, { onConflict: "confirm_date,dept" }).select("confirmed_by, confirmed_at").single();
    if (error) {
      alert(`บันทึกไม่สำเร็จ: ${error.message}\n\nกรุณา run migration 008_daily_confirmations.sql ใน Supabase SQL Editor ก่อน`);
    } else if (data) {
      setConfirmation(data as { confirmed_by: string; confirmed_at: string });
    }
    setIsConfirming(false);
  };

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
      {isoTargetDate && (
        <section className="confirm-attendance-card">
          <div className="confirm-card-top">
            <div className="confirm-card-title-row">
              <ClipboardCheck size={15} className="confirm-title-icon" />
              <span className="confirm-title-text">ยืนยันตรวจสอบการเข้างาน</span>
              <span className="confirm-title-dept">{dashboardDeptFilter === "all" ? "ทุกหน่วยงาน" : dashboardDeptFilter}</span>
            </div>
            <div className="confirm-summary-row">
              <span className="csb present">ตรงเวลา {present}</span>
              <span className="csb late">มาสาย {late}</span>
              <span className="csb absent">ขาด/ลา {absent}</span>
              {absent > 0 && (() => {
                const deptAbsentIds = new Set(allRecords.filter(r => r.status === "Absent").map(r => r.empId));
                return (["ลาป่วย", "ลากิจ", "ลาพักร้อน", "ขาดงาน"] as const).map(type => {
                  const cnt = [...leaveMap.entries()].filter(([eid, lt]) => deptAbsentIds.has(eid) && lt === type).length;
                  return cnt ? <span key={type} className={`csb-leave ${type === "ขาดงาน" ? "red" : type === "ลาพักร้อน" ? "green" : type === "ลากิจ" ? "amber" : "blue"}`}>{type} {cnt}</span> : null;
                });
              })()}
            </div>
          </div>
          <div className="confirm-card-action">
            {confirmation === undefined ? (
              <span className="confirm-loading">กำลังโหลด...</span>
            ) : confirmation ? (
              <div className="confirm-done-row">
                <CheckCircle2 size={15} className="confirm-done-icon" />
                <span>รับทราบแล้วโดย <strong>{confirmation.confirmed_by}</strong> · {new Date(confirmation.confirmed_at).toLocaleString("th-TH", { dateStyle: "short", timeStyle: "short" })}</span>
                <button className="ghost-button small" onClick={() => setConfirmation(null)}>แก้ไข</button>
              </div>
            ) : (
              <div className="confirm-form-row">
                <div className="confirm-input-wrap">
                  <input
                    className="confirm-name-input"
                    placeholder="กรอกชื่อหัวหน้าหน่วยงานก่อนยืนยัน"
                    value={confirmName}
                    onChange={(e) => setConfirmName(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter") handleConfirm(); }}
                  />
                </div>
                <button
                  className={`confirm-submit-btn${!confirmName.trim() ? " needs-name" : ""}`}
                  disabled={isConfirming}
                  onClick={handleConfirm}
                >
                  <CheckCircle2 size={13} />
                  {isConfirming ? "กำลังบันทึก..." : "ยืนยันรับทราบ"}
                </button>
              </div>
            )}
          </div>
        </section>
      )}
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
                  <th>เตือนสะสม</th>
                  <th>เตือน</th>
                </tr>
              </thead>
              <tbody>
                {dashboardLateRows.map((row) => {
                  const warned = warnedIds.has(row.empId);
                  const pending = warnPending.has(row.empId);
                  const warnCount = warnCountMap[row.empId] ?? 0;
                  const riskLevel = warnCount >= 5 ? "fire" : warnCount >= 3 ? "warn" : "";
                  return (
                  <tr key={`dashboard-late-${row.empId}-${row.scanIn}`} className={warned ? "row-warned" : ""}>
                    <td>{row.name}</td>
                    <td><span className="dept-chip">{row.dept}</span></td>
                    <td>{row.scanIn}</td>
                    <td><span className="late-minutes-badge">{formatLateTime(row.minutesLate)}</span></td>
                    <td>
                      <span className={`monthly-count-badge${riskLevel ? ` ${riskLevel}` : ""}`}>
                        {warnCount} ครั้ง
                        {riskLevel === "fire" ? " 🔴" : riskLevel === "warn" ? " 🟡" : ""}
                      </span>
                    </td>
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
                  <tr><td colSpan={6}>ยังไม่มีข้อมูลคนมาสาย</td></tr>
                ) : null}
              </tbody>
            </table>
          </div>
        </section>

        <div className="dashboard-right-col">
        <DonutKpiCard
          present={reportData?.present ?? 0}
          late={reportData?.late ?? 0}
          absent={reportData?.absent ?? 0}
          dayoff={reportData?.dayoff ?? 0}
          total={reportData?.totalEmployees ?? 0}
          totalActive={totalActivePeople}
        />
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
        </div>

      </section>

      {/* Supervisor check panel */}
      {(() => {
        if (dashboardLateRows.length === 0) return null;
        type DeptStat = { late: number; warned: number };
        const deptMap = new Map<string, DeptStat>();
        for (const row of dashboardLateRows) {
          const s = deptMap.get(row.dept) ?? { late: 0, warned: 0 };
          s.late += 1;
          if (warnedIds.has(row.empId)) s.warned += 1;
          deptMap.set(row.dept, s);
        }
        const deptRows = Array.from(deptMap.entries())
          .map(([dept, s]) => ({ dept, ...s, pending: s.late - s.warned }))
          .sort((a, b) => b.pending - a.pending);
        const totalLate = dashboardLateRows.length;
        const totalWarned = dashboardLateRows.filter((r) => warnedIds.has(r.empId)).length;
        const totalPending = totalLate - totalWarned;
        const pct = totalLate ? Math.round((totalWarned / totalLate) * 100) : 0;
        const allDone = totalPending === 0;
        return (
          <section className="panel sup-check-panel">
            <div className="panel-collapse-trigger" onClick={() => setWarnPanelCollapsed(c => !c)}>
              <h3>สถานะการตักเตือน</h3>
              <ChevronDown size={16} className="panel-collapse-chevron" style={{ transform: warnPanelCollapsed ? "rotate(-90deg)" : "rotate(0deg)" }} />
            </div>
            {!warnPanelCollapsed && (
              <>
                <div className="sup-check-header" style={{ marginTop: "12px" }}>
                  <div className="sup-check-title">
                    <p>ตรวจสอบว่าหัวหน้าได้ตักเตือนพนักงานมาสายครบทุกคนแล้วหรือยัง</p>
                  </div>
                  <div className="sup-check-summary">
                    <div className="sup-check-stat">
                      <span className="sup-stat-value">{totalWarned}</span>
                      <span className="sup-stat-label">เตือนแล้ว</span>
                    </div>
                    <div className="sup-check-divider" />
                    <div className="sup-check-stat">
                      <span className="sup-stat-value muted">{totalLate}</span>
                      <span className="sup-stat-label">ทั้งหมด</span>
                    </div>
                    <div className="sup-check-divider" />
                    <div className="sup-check-stat">
                      <span className={`sup-stat-value ${totalPending > 0 ? "danger" : "success"}`}>{totalPending}</span>
                      <span className="sup-stat-label">ยังค้าง</span>
                    </div>
                    <span className={`sup-overall-badge ${allDone ? "done" : "pending"}`}>
                      {allDone ? "✓ ครบแล้ว" : `⚠ ยังค้าง ${totalPending} คน`}
                    </span>
                  </div>
                </div>
                <div className="sup-progress-wrap">
                  <div className="sup-progress-bar">
                    <div className="sup-progress-fill" style={{ width: `${pct}%` }} />
                  </div>
                  <span className="sup-progress-pct">{pct}%</span>
                </div>
                <div className="sup-check-table-wrap">
                  <table className="table compact-table">
                    <thead>
                      <tr>
                        <th>หน่วยงาน</th>
                        <th style={{ textAlign: "center" }}>มาสาย</th>
                        <th style={{ textAlign: "center" }}>เตือนแล้ว</th>
                        <th style={{ textAlign: "center" }}>ยังค้าง</th>
                        <th>ความคืบหน้า</th>
                        <th style={{ textAlign: "center" }}>สถานะ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {deptRows.map((row) => {
                        const deptPct = row.late ? Math.round((row.warned / row.late) * 100) : 100;
                        const done = row.pending === 0;
                        return (
                          <tr key={row.dept} className={done ? "sup-row-done" : "sup-row-pending"}>
                            <td><span className="dept-chip">{row.dept}</span></td>
                            <td style={{ textAlign: "center" }}>{row.late}</td>
                            <td style={{ textAlign: "center" }}><strong style={{ color: "#10b981" }}>{row.warned}</strong></td>
                            <td style={{ textAlign: "center" }}>
                              {row.pending > 0
                                ? <strong style={{ color: "#ef4444" }}>{row.pending}</strong>
                                : <span style={{ color: "#94a3b8" }}>—</span>}
                            </td>
                            <td>
                              <div className="sup-mini-bar-wrap">
                                <div className="sup-mini-bar">
                                  <div className="sup-mini-fill" style={{ width: `${deptPct}%` }} />
                                </div>
                                <span className="sup-mini-pct">{deptPct}%</span>
                              </div>
                            </td>
                            <td style={{ textAlign: "center" }}>
                              <span className={`sup-status-badge ${done ? "done" : "pending"}`}>
                                {done ? "✓ ครบ" : `ค้าง ${row.pending}`}
                              </span>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </>
            )}
          </section>
        );
      })()}

      {/* Employee detail table */}
      <section className="panel detail-attendance-panel">
        <div className="panel-collapse-trigger" onClick={() => setDetailPanelCollapsed(c => !c)}>
          <h3>
            สถานะพนักงานรายคน
            {dashboardDeptFilter !== "all" ? ` · ${dashboardDeptFilter}` : ""}
          </h3>
          <ChevronDown size={16} className="panel-collapse-chevron" style={{ transform: detailPanelCollapsed ? "rotate(-90deg)" : "rotate(0deg)" }} />
        </div>
        {!detailPanelCollapsed && (<>
        <div className="panel-title-row" style={{ marginTop: "12px" }}>
          <div />
          <div className="table-actions">
            <select
              aria-label="กรองสถานะ"
              value={detailStatusFilter}
              onChange={(e) => setDetailStatusFilter(e.target.value)}
            >
              <option value="all">ทุกสถานะ</option>
              <option value="Absent">ขาดงาน</option>
              <option value="Late">มาสาย</option>
              <option value="Present">ตรงเวลา</option>
            </select>
            <button
              className="primary-button small"
              disabled={lateAbsentCount === 0}
              onClick={() => exportLateAbsentToExcel(allRecords, monthlyLateCounts, dashboardDeptFilter === "all" ? "ทุกหน่วยงาน" : dashboardDeptFilter)}
              type="button"
            >
              <Download size={14} />
              Export สาย/ขาด
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
                    <td>
                      {row.status === "Absent" ? (() => {
                        const lt = leaveMap.get(row.empId) ?? "ขาดงาน";
                        const lc = ({ "ลาป่วย": "leave-sick", "ลากิจ": "leave-personal", "ลาพักร้อน": "leave-vacation" } as Record<string, string>)[lt] ?? "leave-absent";
                        return (
                          <select
                            className={`leave-select ${lc}`}
                            value={lt}
                            onChange={(e) => saveLeave(row.empId, e.target.value)}
                          >
                            <option value="ขาดงาน">ขาดงาน</option>
                            <option value="ลาป่วย">ลาป่วย</option>
                            <option value="ลากิจ">ลากิจ</option>
                            <option value="ลาพักร้อน">ลาพักร้อน</option>
                          </select>
                        );
                      })() : (
                        <span className={`status-pill ${row.status.toLowerCase()}`}>{STATUS_TH[row.status] ?? row.status}</span>
                      )}
                    </td>
                    <td>{row.status !== "Absent" && row.status !== "DayOff" ? formatLateTime(row.minutesLate) : "-"}</td>
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
        </>)}
      </section>

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
  onHolidaysChanged,
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
  onHolidaysChanged: (dates: Set<string>) => void;
  saveDayoffShiftRows: (rows: DayoffShiftEditorRow[]) => Promise<void>;
  saveMasterFiles: () => Promise<void>;
  setMasterUploads: Dispatch<SetStateAction<MasterUploadState>>;
}) {
  const [masterSubTab, setMasterSubTab] = useState<"files" | "holidays">("files");

  return (
    <section className="md-page">
      <div className="master-sub-tabs">
        <button
          className={`master-sub-tab${masterSubTab === "files" ? " active" : ""}`}
          onClick={() => setMasterSubTab("files")}
        >
          <FileSpreadsheet size={15} />
          Master Files
        </button>
        <button
          className={`master-sub-tab${masterSubTab === "holidays" ? " active" : ""}`}
          onClick={() => setMasterSubTab("holidays")}
        >
          <CalendarDays size={15} />
          วันพระ
        </button>
      </div>

      {masterSubTab === "holidays" ? (
        <HolidayMasterPage onHolidaysChanged={onHolidaysChanged} />
      ) : null}

      {masterSubTab === "files" ? (<>
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
        employeeMasterFile={activeMasterMap.employee_master}
        saveDayoffShiftRows={saveDayoffShiftRows}
      />
      </>) : null}
    </section>
  );
}

function CalendarPicker({
  value,
  availableDates,
  onChange,
}: {
  value: string;
  availableDates: Map<string, AllocationRun>;
  onChange: (date: string) => void;
}) {
  const [view, setView] = useState<{ year: number; month: number }>(() => {
    const d = value ? new Date(value + "T00:00:00") : new Date();
    return { year: d.getFullYear(), month: d.getMonth() };
  });

  const prevMonth = () => setView(v => v.month === 0 ? { year: v.year - 1, month: 11 } : { ...v, month: v.month - 1 });
  const nextMonth = () => setView(v => v.month === 11 ? { year: v.year + 1, month: 0 } : { ...v, month: v.month + 1 });

  const monthLabel = new Date(view.year, view.month, 1)
    .toLocaleDateString("th-TH", { month: "long", year: "numeric" });

  const firstDay = new Date(view.year, view.month, 1).getDay();
  const daysInMonth = new Date(view.year, view.month + 1, 0).getDate();

  const cells: Array<{ day: number; dateStr: string; hasData: boolean } | null> = [];
  for (let i = 0; i < firstDay; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${view.year}-${String(view.month + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    cells.push({ day: d, dateStr, hasData: availableDates.has(dateStr) });
  }

  return (
    <div className="cal-picker">
      <div className="cal-picker-header">
        <button className="cal-nav-btn" onClick={prevMonth}>‹</button>
        <span className="cal-month-label">{monthLabel}</span>
        <button className="cal-nav-btn" onClick={nextMonth}>›</button>
      </div>
      <div className="cal-weekdays">
        {["อา", "จ", "อ", "พ", "พฤ", "ศ", "ส"].map((d) => (
          <span key={d}>{d}</span>
        ))}
      </div>
      <div className="cal-grid">
        {cells.map((cell, i) =>
          cell === null ? (
            <span key={`e-${i}`} />
          ) : (
            <button
              key={cell.dateStr}
              className={`cal-day${cell.hasData ? "" : " no-data"}${cell.dateStr === value ? " selected" : ""}`}
              onClick={() => cell.hasData && onChange(cell.dateStr)}
              disabled={!cell.hasData}
              title={cell.hasData ? cell.dateStr : undefined}
            >
              {cell.day}
            </button>
          )
        )}
      </div>
    </div>
  );
}

function HolidayMasterPage({
  onHolidaysChanged,
}: {
  onHolidaysChanged: (dates: Set<string>) => void;
}) {
  const [holidays, setHolidays] = useState<HolidayRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [newDate, setNewDate] = useState("");
  const [newName, setNewName] = useState("");
  const [newType, setNewType] = useState<HolidayRow["type"]>("buddhist_holy_day");
  const [yearFilter, setYearFilter] = useState(String(new Date().getFullYear()));
  const [seedYear, setSeedYear] = useState<string | null>(null);

  useEffect(() => {
    void loadHolidays();
  }, []);

  async function loadHolidays() {
    setLoading(true);
    const { data } = await supabase.from("holidays").select("*").order("date");
    if (data) {
      setHolidays(data as HolidayRow[]);
      onHolidaysChanged(new Set(data.map((r: HolidayRow) => r.date)));
    }
    setLoading(false);
  }

  async function addHoliday() {
    if (!newDate || !newName.trim()) return;
    setSaving(true);
    const { error } = await supabase
      .from("holidays")
      .upsert({ date: newDate, name: newName.trim(), type: newType }, { onConflict: "date" });
    if (!error) {
      setNewDate("");
      setNewName("");
      await loadHolidays();
    }
    setSaving(false);
  }

  async function deleteHoliday(id: string) {
    await supabase.from("holidays").delete().eq("id", id);
    await loadHolidays();
  }

  async function seedBuddhistHolyDays(year: string) {
    const dates = buddhistHolyDaysByYear[year];
    if (!dates) return;
    setSeedYear(year);
    const existingDates = new Set(holidays.map((h) => h.date));
    const toInsert = Array.from(dates)
      .filter((d) => !existingDates.has(d))
      .map((d) => ({ date: d, name: "วันพระ", type: "buddhist_holy_day" as const }));
    if (toInsert.length > 0) {
      await supabase.from("holidays").insert(toInsert);
    }
    await loadHolidays();
    setSeedYear(null);
  }

  const existingYears = Array.from(new Set(holidays.map((h) => h.date.substring(0, 4)))).sort();
  const allYears = Array.from(new Set([...existingYears, yearFilter])).sort();
  const seedableYears = Object.keys(buddhistHolyDaysByYear).sort();
  const filtered = holidays.filter((h) => h.date.startsWith(yearFilter));

  const typeLabel: Record<HolidayRow["type"], string> = {
    buddhist_holy_day: "วันพระ",
    public_holiday: "วันหยุดราชการ",
    company_holiday: "วันหยุดบริษัท",
  };
  const typeBadgeClass: Record<HolidayRow["type"], string> = {
    buddhist_holy_day: "holiday-badge-buddhist",
    public_holiday: "holiday-badge-public",
    company_holiday: "holiday-badge-company",
  };

  return (
    <section className="panel holiday-master-panel">
      <div className="holiday-master-header">
        <div className="holiday-master-title">
          <CalendarDays size={20} />
          <h2>จัดการวันพระ</h2>
          <span className="holiday-total-badge">{holidays.length} วัน</span>
        </div>
        <span style={{ fontSize: 13, color: "var(--muted)" }}>
          ข้อมูลอัพเดทถึงปี {seedableYears[seedableYears.length - 1]}
        </span>
      </div>

      <div className="holiday-add-form">
        <div className="holiday-add-form-fields">
          <input
            type="date"
            className="holiday-input"
            value={newDate}
            onChange={(e) => setNewDate(e.target.value)}
          />
          <input
            type="text"
            className="holiday-input holiday-input-name"
            placeholder="ชื่อวันหยุด เช่น วันสงกรานต์"
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && void addHoliday()}
          />
          <select
            className="holiday-type-select"
            value={newType}
            onChange={(e) => setNewType(e.target.value as HolidayRow["type"])}
          >
            <option value="buddhist_holy_day">วันพระ</option>
            <option value="public_holiday">วันหยุดราชการ</option>
            <option value="company_holiday">วันหยุดบริษัท</option>
          </select>
          <button
            className="primary-button"
            onClick={addHoliday}
            disabled={saving || !newDate || !newName.trim()}
          >
            + เพิ่มวันหยุด
          </button>
        </div>
      </div>

      <div className="holiday-year-tabs">
        {allYears.map((yr) => (
          <button
            key={yr}
            className={`holiday-year-tab${yearFilter === yr ? " active" : ""}`}
            onClick={() => setYearFilter(yr)}
          >
            {yr}
            <span className="holiday-year-count">
              {holidays.filter((h) => h.date.startsWith(yr)).length}
            </span>
          </button>
        ))}
      </div>

      <div className="holiday-table-wrap">
        {loading ? (
          <p className="holiday-loading">กำลังโหลด…</p>
        ) : (
          <table className="table holiday-table">
            <thead>
              <tr>
                <th>วันที่</th>
                <th>วันในสัปดาห์</th>
                <th>ชื่อวันหยุด</th>
                <th>ประเภท</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((h) => {
                const d = new Date(h.date + "T00:00:00");
                return (
                  <tr key={h.id}>
                    <td className="holiday-date-cell">
                      {d.toLocaleDateString("th-TH", { day: "numeric", month: "long", year: "numeric" })}
                    </td>
                    <td>
                      {["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"][d.getDay()]}
                    </td>
                    <td>{h.name}</td>
                    <td>
                      <span className={`holiday-type-badge ${typeBadgeClass[h.type]}`}>
                        {typeLabel[h.type]}
                      </span>
                    </td>
                    <td>
                      <button
                        className="holiday-delete-btn"
                        onClick={() => deleteHoliday(h.id)}
                        title="ลบวันหยุดนี้"
                      >
                        <X size={14} />
                      </button>
                    </td>
                  </tr>
                );
              })}
              {filtered.length === 0 && (
                <tr>
                  <td colSpan={5} className="holiday-empty-row">
                    ไม่มีวันหยุดสำหรับปี {yearFilter} — กด Seed หรือเพิ่มเองด้านบน
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        )}
      </div>
    </section>
  );
}

function DayoffShiftEditor({
  activeFile,
  employeeMasterFile,
  saveDayoffShiftRows,
}: {
  activeFile?: MasterFile;
  employeeMasterFile?: MasterFile;
  saveDayoffShiftRows: (rows: DayoffShiftEditorRow[]) => Promise<void>;
}) {
  const [rows, setRows] = useState<DayoffShiftEditorRow[]>([]);
  const [originalRows, setOriginalRows] = useState<DayoffShiftEditorRow[]>([]);
  const [query, setQuery] = useState("");
  const [selectedDept, setSelectedDept] = useState("all");
  const [selectedDayoff, setSelectedDayoff] = useState("all");
  const [selectedShift, setSelectedShift] = useState("all");
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [bulkDayoff, setBulkDayoff] = useState("");
  const [bulkShift, setBulkShift] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  const dayoffSingle = [
    { value: "จ",  label: "จ — จันทร์" },
    { value: "อ",  label: "อ — อังคาร" },
    { value: "พ",  label: "พ — พุธ" },
    { value: "พฤ", label: "พฤ — พฤหัส" },
    { value: "ศ",  label: "ศ — ศุกร์" },
    { value: "ส",  label: "ส — เสาร์" },
    { value: "อา", label: "อา — อาทิตย์" },
    { value: "พระ",label: "พระ — วันพระ" },
  ];
  const dayoffDouble = [
    { value: "ส,อา", label: "ส+อา — เสาร์&อาทิตย์" },
  ];
  const shiftOptions = Array.from(new Set([
    "กะ1", "กะ2", "กะ3",
    ...rows.map((r) => r.shift).filter(Boolean),
  ]));
  const deptOptions = Array.from(new Set(rows.map((r) => r.dept).filter(Boolean))).sort();

  const normalizedQuery = query.trim().toLowerCase();
  const filteredRows = rows.filter((row) => {
    if (selectedDept !== "all" && row.dept !== selectedDept) return false;
    if (selectedDayoff !== "all") {
      if (selectedDayoff === "__empty__" && row.dayoff !== "") return false;
      if (selectedDayoff !== "__empty__" && row.dayoff !== selectedDayoff) return false;
    }
    if (selectedShift !== "all") {
      if (selectedShift === "__empty__" && row.shift !== "") return false;
      if (selectedShift !== "__empty__" && row.shift !== selectedShift) return false;
    }
    if (!normalizedQuery) return true;
    return [row.empId, row.name, row.dept, row.dayoff, row.shift]
      .some((v) => v.toLowerCase().includes(normalizedQuery));
  });

  const modifiedIds = new Set(
    rows
      .filter((row) => {
        const orig = originalRows.find((r) => r.id === row.id);
        return orig && (orig.dayoff !== row.dayoff || orig.shift !== row.shift || orig.shiftStart !== row.shiftStart);
      })
      .map((r) => r.id),
  );

  const allFilteredSelected = filteredRows.length > 0 && filteredRows.every((r) => selectedIds.has(r.id));

  useEffect(() => {
    if (!activeFile?.file_path) {
      setRows([]);
      setOriginalRows([]);
      setSelectedIds(new Set());
      return;
    }

    let isMounted = true;
    setIsLoading(true);

    // กะ1=AI(col 34), กะ2=AL(col 37), กะ3=AO(col 40)
    const SHIFT_TIME_COL: Record<string, number> = { "กะ1": 34, "กะ2": 37, "กะ3": 40 };

    const dayoffPromise = downloadSheetRows(activeFile.file_path);
    const empPromise: Promise<{ rows: Record<string, unknown>[]; rawRows: unknown[][] }> =
      employeeMasterFile?.file_path
        ? (async () => {
            const { data, error } = await supabase.storage.from("workforce-inputs").download(employeeMasterFile.file_path);
            if (error || !data) return { rows: [], rawRows: [] };
            const buffer = await data.arrayBuffer();
            const wb = XLSX.read(buffer, { type: "array", cellDates: true });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            return {
              rows: XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "", raw: true }),
              rawRows: XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true }),
            };
          })()
        : Promise.resolve({ rows: [] as Record<string, unknown>[], rawRows: [] as unknown[][] });

    Promise.all([dayoffPromise, empPromise])
      .then(([dayoffRows, { rows: empRows, rawRows: empRawRows }]) => {
        if (!isMounted) return;
        const deptMap = new Map<string, string>();
        const timeMap = new Map<string, string>();
        for (const row of empRows) {
          const empId = cleanEmpId(
            row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"],
          );
          const dept = String(
            row["หน่วยงาน"] ?? row["Name (Section)"] ?? row["Department"] ?? "",
          ).trim();
          if (empId && dept) deptMap.set(empId, dept);
        }
        // ดึงเวลาเข้าจาก column index ตาม shift (AI=กะ1, AL=กะ2, AO=กะ3)
        const hdrIdx = empRawRows.findIndex((r) =>
          (r as unknown[]).some((c) =>
            ["User ID (Job Information)", "Employee ID", "Emp ID"].includes(String(c))
          )
        );
        if (hdrIdx >= 0) {
          const hdr = empRawRows[hdrIdx] as unknown[];
          const empColIdx = hdr.findIndex((c) =>
            ["User ID (Job Information)", "Employee ID", "Emp ID"].includes(String(c))
          );
          const shiftColIdx = hdr.findIndex((c) =>
            ["อยู่กะไหน", "shift", "กะ", "Shift"].includes(String(c))
          );
          for (const rawRow of empRawRows.slice(hdrIdx + 1)) {
            const r = rawRow as unknown[];
            const empId = empColIdx >= 0 ? cleanEmpId(r[empColIdx]) : "";
            if (!empId) continue;
            const shiftKey = normalizeShiftKey(shiftColIdx >= 0 ? r[shiftColIdx] : "กะ1");
            const colIdx = SHIFT_TIME_COL[shiftKey] ?? SHIFT_TIME_COL["กะ1"];
            const t = normalizeTimeText(r[colIdx]);
            if (t) timeMap.set(empId, t);
          }
        }
        const parsed = dayoffRows.map((row, i) => {
          const r = toDayoffShiftEditorRow(row, i);
          if (!r.dept && deptMap.has(r.empId)) r.dept = deptMap.get(r.empId)!;
          if (!r.shiftStart && timeMap.has(r.empId)) r.shiftStart = timeMap.get(r.empId)!;
          return r;
        });
        setRows(parsed);
        setOriginalRows(parsed);
        setSelectedIds(new Set());
      })
      .catch(() => {
        if (!isMounted) return;
        setRows([]);
        setOriginalRows([]);
      })
      .finally(() => {
        if (!isMounted) return;
        setIsLoading(false);
      });

    return () => { isMounted = false; };
  }, [activeFile?.file_path, employeeMasterFile?.file_path]);

  function updateRow(id: string, field: "dayoff" | "shift" | "shiftStart", value: string) {
    const rawKey = field === "dayoff" ? "วันหยุด\nประจำสัปดาห์" : field === "shift" ? "อยู่กะไหน" : "เวลาเข้างาน";
    setRows((current) =>
      current.map((row) =>
        row.id === id
          ? { ...row, [field]: value, raw: { ...row.raw, [rawKey]: value } }
          : row,
      ),
    );
  }

  function toggleSelect(id: string) {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    if (allFilteredSelected) {
      setSelectedIds((prev) => {
        const next = new Set(prev);
        filteredRows.forEach((r) => next.delete(r.id));
        return next;
      });
    } else {
      setSelectedIds((prev) => {
        const next = new Set(prev);
        filteredRows.forEach((r) => next.add(r.id));
        return next;
      });
    }
  }

  function applyBulk() {
    if (!bulkDayoff && !bulkShift) return;
    setRows((current) =>
      current.map((row) => {
        if (!selectedIds.has(row.id)) return row;
        return {
          ...row,
          dayoff: bulkDayoff || row.dayoff,
          shift: bulkShift || row.shift,
          raw: {
            ...row.raw,
            ...(bulkDayoff ? { "วันหยุด\nประจำสัปดาห์": bulkDayoff } : {}),
            ...(bulkShift ? { "อยู่กะไหน": bulkShift } : {}),
          },
        };
      }),
    );
    setSelectedIds(new Set());
    setBulkDayoff("");
    setBulkShift("");
  }

  async function handleSave() {
    setIsSaving(true);
    try {
      await saveDayoffShiftRows(rows);
      setOriginalRows(rows);
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
          {isSaving ? "Saving..." : `Save${modifiedIds.size > 0 ? ` (${modifiedIds.size} แก้ไข)` : ""}`}
        </button>
      </div>

      <div className="table-filters dayoff-editor-filters">
        <input
          aria-label="ค้นหา dayoff shift"
          placeholder="ค้นหา รหัส ชื่อ แผนก"
          type="search"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
        <select value={selectedDept} onChange={(e) => setSelectedDept(e.target.value)}>
          <option value="all">ทุกแผนก</option>
          {deptOptions.map((d) => <option key={d} value={d}>{d}</option>)}
        </select>
        <select value={selectedDayoff} onChange={(e) => setSelectedDayoff(e.target.value)}>
          <option value="all">ทุก Dayoff</option>
          <option value="__empty__">— ยังไม่ได้ตั้ง</option>
          <optgroup label="หยุด 1 วัน">
            {dayoffSingle.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
          </optgroup>
          <optgroup label="หยุด 2 วัน">
            {dayoffDouble.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
          </optgroup>
        </select>
        <select value={selectedShift} onChange={(e) => setSelectedShift(e.target.value)}>
          <option value="all">ทุก Shift</option>
          <option value="__empty__">— ยังไม่ได้ตั้ง</option>
          {shiftOptions.map((o) => <option key={o} value={o}>{o}</option>)}
        </select>
        <span className="dayoff-count">
          {filteredRows.length.toLocaleString()} / {rows.length.toLocaleString()} คน
          {modifiedIds.size > 0 && <span className="modified-badge">{modifiedIds.size} แก้ไข</span>}
        </span>
      </div>

      {selectedIds.size > 0 && (
        <div className="dayoff-bulk-bar">
          <span className="bulk-count">{selectedIds.size} คนที่เลือก</span>
          <select value={bulkDayoff} onChange={(e) => setBulkDayoff(e.target.value)}>
            <option value="">เปลี่ยน Dayoff...</option>
            <optgroup label="หยุด 1 วัน">
              {dayoffSingle.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
            </optgroup>
            <optgroup label="หยุด 2 วัน">
              {dayoffDouble.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
            </optgroup>
          </select>
          <select value={bulkShift} onChange={(e) => setBulkShift(e.target.value)}>
            <option value="">เปลี่ยน Shift...</option>
            {shiftOptions.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
          <button
            className="primary-button"
            disabled={!bulkDayoff && !bulkShift}
            onClick={applyBulk}
            style={{ height: 32, fontSize: 12, padding: "0 14px" }}
            type="button"
          >
            Apply
          </button>
          <button
            className="secondary-button"
            onClick={() => setSelectedIds(new Set())}
            style={{ height: 32, fontSize: 12, padding: "0 14px" }}
            type="button"
          >
            ยกเลิก
          </button>
        </div>
      )}

      <div className="dayoff-editor-table">
        <table className="table">
          <thead>
            <tr>
              <th style={{ width: 36 }}>
                <input
                  type="checkbox"
                  checked={allFilteredSelected}
                  onChange={toggleSelectAll}
                  title="เลือกทั้งหมดในมุมมองนี้"
                />
              </th>
              <th>Emp ID</th>
              <th>ชื่อ</th>
              <th>แผนก</th>
              <th>Dayoff</th>
              <th>Shift</th>
              <th>เวลาเข้า</th>
              <th>เวลาออก</th>
            </tr>
          </thead>
          <tbody>
            {filteredRows.map((row) => (
              <tr
                key={row.id}
                className={[
                  selectedIds.has(row.id) ? "row-selected" : "",
                  modifiedIds.has(row.id) ? "row-modified" : "",
                ].filter(Boolean).join(" ")}
              >
                <td>
                  <input
                    type="checkbox"
                    checked={selectedIds.has(row.id)}
                    onChange={() => toggleSelect(row.id)}
                  />
                </td>
                <td>{row.empId}</td>
                <td>{row.name}</td>
                <td className="dept-cell">{row.dept || "—"}</td>
                <td>
                  <select value={row.dayoff} onChange={(e) => updateRow(row.id, "dayoff", e.target.value)}>
                    <option value="">-</option>
                    <optgroup label="หยุด 1 วัน">
                      {dayoffSingle.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
                    </optgroup>
                    <optgroup label="หยุด 2 วัน">
                      {dayoffDouble.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
                    </optgroup>
                  </select>
                </td>
                <td>
                  <select value={row.shift} onChange={(e) => updateRow(row.id, "shift", e.target.value)}>
                    <option value="">-</option>
                    {shiftOptions.map((o) => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>
                <td>
                  <input
                    className="shift-time-input"
                    type="time"
                    value={row.shiftStart}
                    onChange={(e) => updateRow(row.id, "shiftStart", e.target.value)}
                  />
                </td>
                <td className="shift-end-cell">
                  {row.shiftStart ? addHoursToTime(row.shiftStart, 9) : "—"}
                </td>
              </tr>
            ))}
            {!activeFile ? (
              <tr><td colSpan={8}>อัปโหลด Dayoff & Shift master ก่อน จึงจะแก้ไขในหน้านี้ได้</td></tr>
            ) : null}
            {activeFile && isLoading ? (
              <tr><td colSpan={8}>Loading Dayoff & Shift...</td></tr>
            ) : null}
            {activeFile && !isLoading && filteredRows.length === 0 ? (
              <tr><td colSpan={8}>ไม่พบข้อมูลที่ค้นหา</td></tr>
            ) : null}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function SkillMatrixPage({
  activeFile,
  employeeMasterFile,
  saveSkillMatrixRows,
}: {
  activeFile?: MasterFile;
  employeeMasterFile?: MasterFile;
  saveSkillMatrixRows: (rows: SkillMatrixSaveRow[]) => Promise<void>;
}) {
  const [rows, setRows] = useState<SkillFlatRow[]>([]);
  const [empInfoMap, setEmpInfoMap] = useState<Map<string, { name: string; dept: string }>>(new Map());
  const [query, setQuery] = useState("");
  const [selectedDept, setSelectedDept] = useState("all");
  const [selectedSkill, setSelectedSkill] = useState("all");
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [bulkLevel, setBulkLevel] = useState("");
  const [addEmpId, setAddEmpId] = useState("");
  const [addSkill, setAddSkill] = useState("");
  const [addLevel, setAddLevel] = useState(1);
  const [addOpen, setAddOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  useEffect(() => {
    if (!activeFile?.file_path) {
      setRows([]);
      setSelectedIds(new Set());
      return;
    }

    let isMounted = true;
    setIsLoading(true);

    const skillPromise = downloadSheetRows(activeFile.file_path);
    const empPromise = employeeMasterFile?.file_path
      ? downloadSheetRows(employeeMasterFile.file_path)
      : Promise.resolve([] as Record<string, unknown>[]);

    Promise.all([skillPromise, empPromise])
      .then(([skillRows, empRows]) => {
        if (!isMounted) return;

        const empInfo = new Map<string, { name: string; dept: string }>();
        for (const row of empRows) {
          const empId = cleanEmpId(
            row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"],
          );
          const firstName = String(row["First Name (Local)"] ?? "").trim();
          const lastName = String(row["Last Name (Local)"] ?? "").trim();
          const name =
            `${firstName} ${lastName}`.trim() ||
            String(row["Employee Name"] ?? "").trim() ||
            empId;
          const dept = String(row["หน่วยงาน"] ?? row["Name (Section)"] ?? "").trim();
          if (empId) empInfo.set(empId, { name, dept });
        }

        const parsed: SkillFlatRow[] = skillRows
          .map((row, i) => {
            const empId = cleanEmpId(
              row["Employee ID"] ?? row["Emp ID"] ?? row["emp_id"] ?? row["รหัสพนักงาน"],
            );
            const skill = String(row["Skill"] ?? row["skill"] ?? row["ทักษะ"] ?? "").trim();
            const level = Number(row["Level"] ?? row["level"] ?? row["ระดับ"]) || 0;
            const info = empInfo.get(empId) ?? { name: empId, dept: "" };
            return {
              id: `${i}-${empId}-${skill}`,
              empId,
              name: info.name,
              dept: info.dept,
              skill,
              level,
              origLevel: level,
            };
          })
          .filter((r) => r.empId && r.skill);

        setRows(parsed);
        setEmpInfoMap(empInfo);
        setSelectedIds(new Set());
      })
      .catch(() => {
        if (!isMounted) return;
        setRows([]);
      })
      .finally(() => {
        if (!isMounted) return;
        setIsLoading(false);
      });

    return () => { isMounted = false; };
  }, [activeFile?.file_path, employeeMasterFile?.file_path]);

  const deptOptions = Array.from(new Set(rows.map((r) => r.dept).filter(Boolean))).sort();
  const skillOptions = Array.from(new Set(rows.map((r) => r.skill).filter(Boolean))).sort();
  const normalizedQuery = query.trim().toLowerCase();
  const filteredRows = rows.filter((row) => {
    if (selectedDept !== "all" && row.dept !== selectedDept) return false;
    if (selectedSkill !== "all" && row.skill !== selectedSkill) return false;
    if (!normalizedQuery) return true;
    return [row.empId, row.name, row.dept, row.skill].some((v) =>
      v.toLowerCase().includes(normalizedQuery),
    );
  });

  const modifiedIds = new Set(
    rows.filter((r) => r.level !== r.origLevel).map((r) => r.id),
  );
  const allFilteredSelected =
    filteredRows.length > 0 && filteredRows.every((r) => selectedIds.has(r.id));

  function updateRow(id: string, level: number) {
    setRows((prev) => prev.map((r) => (r.id === id ? { ...r, level } : r)));
  }

  function toggleSelect(id: string) {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  function toggleSelectAll() {
    if (allFilteredSelected) {
      setSelectedIds((prev) => {
        const next = new Set(prev);
        filteredRows.forEach((r) => next.delete(r.id));
        return next;
      });
    } else {
      setSelectedIds((prev) => {
        const next = new Set(prev);
        filteredRows.forEach((r) => next.add(r.id));
        return next;
      });
    }
  }

  function applyBulk() {
    if (!bulkLevel) return;
    const level = Number(bulkLevel);
    setRows((prev) =>
      prev.map((r) => (selectedIds.has(r.id) ? { ...r, level } : r)),
    );
    setSelectedIds(new Set());
    setBulkLevel("");
  }

  async function handleSave() {
    setIsSaving(true);
    try {
      const flatRows = rows
        .filter((r) => r.level > 0)
        .map((r) => ({ empId: r.empId, skill: r.skill, level: r.level }));
      await saveSkillMatrixRows(flatRows);
      setRows((prev) => prev.map((r) => ({ ...r, origLevel: r.level })));
    } finally {
      setIsSaving(false);
    }
  }

  function addRow() {
    const empId = addEmpId.trim();
    const skill = addSkill.trim();
    if (!empId || !skill) return;

    const existing = rows.find((r) => r.empId === empId && r.skill === skill);
    if (existing) {
      setRows((prev) =>
        prev.map((r) => (r.empId === empId && r.skill === skill ? { ...r, level: addLevel } : r)),
      );
    } else {
      const info = empInfoMap.get(empId) ??
        rows.find((r) => r.empId === empId) ??
        { name: empId, dept: "" };
      setRows((prev) => [
        ...prev,
        {
          id: `add-${Date.now()}-${empId}-${skill}`,
          empId,
          name: info.name,
          dept: info.dept,
          skill,
          level: addLevel,
          origLevel: 0,
        },
      ]);
    }
    setAddEmpId("");
    setAddSkill("");
    setAddLevel(1);
    setAddOpen(false);
  }

  const levelBg: Record<number, string> = {
    1: "#fecaca", 2: "#fed7aa", 3: "#fef08a", 4: "#bbf7d0", 5: "#34d399",
  };

  return (
    <section className="panel dayoff-editor-panel">
      <div className="panel-title-row">
        <div>
          <h3>Skill Matrix</h3>
          <p>แก้ไขระดับทักษะรายคน — กรอง Skill เพื่อแก้หลายคนพร้อมกัน หรือเลือกหลายแถว bulk edit</p>
        </div>
        <button
          className="primary-button"
          disabled={!rows.length || isSaving}
          onClick={handleSave}
          type="button"
        >
          <UploadCloud size={17} />
          {isSaving ? "Saving..." : `Save${modifiedIds.size > 0 ? ` (${modifiedIds.size} แก้ไข)` : ""}`}
        </button>
      </div>

      <div className="table-filters dayoff-editor-filters" style={{ gridTemplateColumns: "1fr 180px 180px auto" }}>
        <input
          aria-label="ค้นหา skill matrix"
          placeholder="ค้นหา รหัส ชื่อ แผนก ทักษะ"
          type="search"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
        <select value={selectedDept} onChange={(e) => setSelectedDept(e.target.value)}>
          <option value="all">ทุกแผนก</option>
          {deptOptions.map((d) => <option key={d} value={d}>{d}</option>)}
        </select>
        <select value={selectedSkill} onChange={(e) => setSelectedSkill(e.target.value)}>
          <option value="all">ทุก Skill</option>
          {skillOptions.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>
        <span className="dayoff-count">
          {filteredRows.length.toLocaleString()} / {rows.length.toLocaleString()} แถว
          {modifiedIds.size > 0 && <span className="modified-badge">{modifiedIds.size} แก้ไข</span>}
        </span>
      </div>

      {selectedIds.size > 0 && (
        <div className="dayoff-bulk-bar">
          <span className="bulk-count">{selectedIds.size} แถวที่เลือก</span>
          <select value={bulkLevel} onChange={(e) => setBulkLevel(e.target.value)}>
            <option value="">เปลี่ยน Level...</option>
            <option value="0">0 — ไม่มี</option>
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
          </select>
          <button
            className="primary-button"
            disabled={!bulkLevel}
            onClick={applyBulk}
            style={{ height: 32, fontSize: 12, padding: "0 14px" }}
            type="button"
          >
            Apply
          </button>
          <button
            className="secondary-button"
            onClick={() => setSelectedIds(new Set())}
            style={{ height: 32, fontSize: 12, padding: "0 14px" }}
            type="button"
          >
            ยกเลิก
          </button>
        </div>
      )}

      <div className="dayoff-editor-table">
        <table className="table">
          <thead>
            <tr>
              <th style={{ width: 36 }}>
                <input
                  type="checkbox"
                  checked={allFilteredSelected}
                  onChange={toggleSelectAll}
                  title="เลือกทั้งหมดในมุมมองนี้"
                />
              </th>
              <th>Emp ID</th>
              <th>ชื่อ</th>
              <th>แผนก</th>
              <th>Skill</th>
              <th>Level</th>
            </tr>
          </thead>
          <tbody>
            {filteredRows.map((row) => (
              <tr
                key={row.id}
                className={[
                  selectedIds.has(row.id) ? "row-selected" : "",
                  modifiedIds.has(row.id) ? "row-modified" : "",
                ].filter(Boolean).join(" ")}
              >
                <td>
                  <input
                    type="checkbox"
                    checked={selectedIds.has(row.id)}
                    onChange={() => toggleSelect(row.id)}
                  />
                </td>
                <td>{row.empId}</td>
                <td>{row.name}</td>
                <td className="dept-cell">{row.dept || "—"}</td>
                <td>{row.skill}</td>
                <td>
                  <select
                    className="skill-level-cell"
                    style={{ background: levelBg[row.level] ?? "" }}
                    value={row.level}
                    onChange={(e) => updateRow(row.id, Number(e.target.value))}
                  >
                    <option value={0}>— (0)</option>
                    <option value={1}>1</option>
                    <option value={2}>2</option>
                    <option value={3}>3</option>
                    <option value={4}>4</option>
                    <option value={5}>5</option>
                  </select>
                </td>
              </tr>
            ))}
            {!activeFile ? (
              <tr><td colSpan={6}>อัปโหลด Skill Matrix master ก่อน จึงจะแก้ไขในหน้านี้ได้</td></tr>
            ) : null}
            {activeFile && isLoading ? (
              <tr><td colSpan={6}>Loading Skill Matrix...</td></tr>
            ) : null}
            {activeFile && !isLoading && filteredRows.length === 0 ? (
              <tr><td colSpan={6}>ไม่พบข้อมูลที่ค้นหา</td></tr>
            ) : null}
          </tbody>
        </table>
      </div>

      {/* Add row */}
      {activeFile && (
        <div className="skill-add-row-bar">
          {addOpen ? (
            <>
              <input
                className="skill-add-empid"
                list="skill-emp-datalist"
                placeholder="Emp ID หรือชื่อ..."
                value={addEmpId}
                onChange={(e) => setAddEmpId(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Escape") setAddOpen(false); }}
              />
              <datalist id="skill-emp-datalist">
                {Array.from(new Set(rows.map((r) => r.empId))).map((id) => {
                  const r = rows.find((x) => x.empId === id);
                  return <option key={id} value={id}>{r?.name ?? id}</option>;
                })}
              </datalist>
              <input
                className="skill-add-skillname"
                list="skill-name-datalist"
                placeholder="ชื่อ Skill..."
                value={addSkill}
                onChange={(e) => setAddSkill(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === "Enter") addRow();
                  if (e.key === "Escape") setAddOpen(false);
                }}
              />
              <datalist id="skill-name-datalist">
                {Array.from(new Set(rows.map((r) => r.skill))).sort().map((s) => (
                  <option key={s} value={s} />
                ))}
              </datalist>
              <select
                value={addLevel}
                onChange={(e) => setAddLevel(Number(e.target.value))}
                style={{ background: levelBg[addLevel] ?? "", height: 34, border: "1px solid var(--line)", borderRadius: 6, padding: "0 8px", fontWeight: 600 }}
              >
                {[1, 2, 3, 4, 5].map((v) => <option key={v} value={v}>{v}</option>)}
              </select>
              <button
                className="primary-button"
                disabled={!addEmpId.trim() || !addSkill.trim()}
                onClick={addRow}
                style={{ height: 34, fontSize: 13, padding: "0 16px" }}
                type="button"
              >
                เพิ่ม
              </button>
              <button
                className="secondary-button"
                onClick={() => { setAddOpen(false); setAddEmpId(""); setAddSkill(""); }}
                style={{ height: 34, fontSize: 13, padding: "0 12px" }}
                type="button"
              >
                ยกเลิก
              </button>
            </>
          ) : (
            <button className="skill-add-row-btn" onClick={() => setAddOpen(true)} type="button">
              + เพิ่ม Skill ให้พนักงาน
            </button>
          )}
        </div>
      )}
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
          <option value="Present">ตรงเวลา</option>
          <option value="Late">มาสาย</option>
          <option value="Absent">ขาดงาน</option>
          <option value="DayOff">วันหยุด</option>
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
                <td><span className={`status-pill ${row.status.toLowerCase()}`}>{STATUS_TH[row.status] ?? row.status}</span></td>
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
            <option value="Present">ตรงเวลา</option>
            <option value="Late">มาสาย</option>
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
                    {row.status === "DayOff" ? "วันหยุด" : row.status}
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
  dailyStats,
  deptFilter,
  monthlyLateMinutes,
  prevMonthLateCounts,
  query,
  reportData,
  selectedDept,
  setDeptFilter,
  setQuery,
  setSelectedDept,
  warnCountMap,
  warnDates,
}: {
  dailyStats: DailyStat[];
  deptFilter: string;
  monthlyLateMinutes: Record<string, number>;
  prevMonthLateCounts: Record<string, number>;
  query: string;
  reportData: ReportData | null;
  selectedDept: string;
  setDeptFilter: (value: string) => void;
  setQuery: (value: string) => void;
  setSelectedDept: (value: string) => void;
  warnCountMap: Record<string, number>;
  warnDates: Record<string, string[]>;
}) {
  const [sort, setSort_] = useState<SortState>(null);
  const setSort = setSort_ as (sort: SortState) => void;
  const [tableStatusFilter, setTableStatusFilter] = useState<"all" | "Late" | "Absent" | "DayOff">("all");
  const [leaveMap, setLeaveMap] = useState<Map<string, string>>(new Map());

  const isoDate = reportData?.isoTargetDate ?? "";
  const [deptConfirmations, setDeptConfirmations] = useState<Map<string, { confirmed_by: string; confirmed_at: string }>>(new Map());
  const [confirmPanelCollapsed, setConfirmPanelCollapsed] = useState(true);
  const [warnPanelCollapsed, setWarnPanelCollapsed] = useState(true);
  const [lateAbsentCollapsed, setLateAbsentCollapsed] = useState(true);

  useEffect(() => {
    if (!isoDate) return;
    supabase.from("leave_records").select("emp_id, leave_type").eq("leave_date", isoDate)
      .then(({ data: rows }) => {
        if (!rows) return;
        setLeaveMap(new Map(rows.map((r: { emp_id: string; leave_type: string }) => [r.emp_id, r.leave_type])));
      });
    supabase.from("daily_confirmations")
      .select("dept, confirmed_by, confirmed_at")
      .eq("confirm_date", isoDate)
      .then(({ data: rows }) => {
        if (!rows) return;
        setDeptConfirmations(new Map(rows.map((r: { dept: string; confirmed_by: string; confirmed_at: string }) => [r.dept, { confirmed_by: r.confirmed_by, confirmed_at: r.confirmed_at }])));
      });
  }, [isoDate]);

  const saveLeave = async (empId: string, leaveType: string) => {
    if (!isoDate || !leaveType) return;
    setLeaveMap(prev => new Map(prev).set(empId, leaveType));
    await supabase.from("leave_records").upsert(
      { emp_id: empId, leave_date: isoDate, leave_type: leaveType },
      { onConflict: "emp_id,leave_date" }
    );
  };

  const data = reportData ?? {
    targetDate: "-",
    totalEmployees: 0,
    present: 0,
    late: 0,
    absent: 0,
    dayoff: 0,
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
  const scopedDayoff = scopedRecords.filter((row) => row.status === "DayOff").length;
  const scopedCameToWork = scopedPresent + scopedLate;
  const activeTotal = scopedTotal - scopedDayoff; // exclude DayOff from % denominator
  const riskCount = new Set(
    scopedRecords.filter(r => (data.monthlyLateCounts[r.empId] ?? 0) >= 3).map(r => r.empId)
  ).size;
  const lateRate = scopedCameToWork
    ? ((scopedLate / scopedCameToWork) * 100).toFixed(1)
    : "0.0";
  const maxDeptTotal = Math.max(...data.deptRows.map((row) => row.present + row.late + row.absent), 1);
  const presentPercent = activeTotal ? (scopedPresent / activeTotal) * 100 : 0;
  const latePercent = activeTotal ? (scopedLate / activeTotal) * 100 : 0;
  const absentPercent = activeTotal ? (scopedAbsent / activeTotal) * 100 : 0;
  const warnedOnDate = new Set(
    Object.entries(warnDates)
      .filter(([, dates]) => data.isoTargetDate ? dates.includes(data.isoTargetDate) : false)
      .map(([id]) => id),
  );
  const scopedLateRows = selectedDept === "all"
    ? data.lateRows
    : data.lateRows.filter((r) => r.dept === selectedDept);
  const supDeptMap = new Map<string, { late: number; warned: number }>();
  for (const row of scopedLateRows) {
    const s = supDeptMap.get(row.dept) ?? { late: 0, warned: 0 };
    s.late += 1;
    if (warnedOnDate.has(row.empId)) s.warned += 1;
    supDeptMap.set(row.dept, s);
  }
  const supDeptRows = Array.from(supDeptMap.entries())
    .map(([dept, s]) => ({ dept, ...s, pending: s.late - s.warned }))
    .sort((a, b) => b.pending - a.pending);
  const supTotalLate = scopedLateRows.length;
  const supTotalWarned = scopedLateRows.filter((r) => warnedOnDate.has(r.empId)).length;
  const supPct = supTotalLate ? Math.round((supTotalWarned / supTotalLate) * 100) : 0;
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
  const tableSourceRows = scopedRecords.filter((r) => r.status === "Late" || r.status === "Absent" || r.status === "DayOff");
  const tableDeptOptions = Array.from(new Set(tableSourceRows.map((r) => r.dept))).sort();
  const filteredTableRows = tableSourceRows.filter((row) => {
    const matchesQuery = !normalizedQuery || [
      row.empId, row.name, row.dept, row.position, row.scanIn, row.status,
    ].some((v) => String(v).toLowerCase().includes(normalizedQuery));
    const matchesDept = deptFilter === "all" || row.dept === deptFilter;
    const matchesStatus = tableStatusFilter === "all" || row.status === tableStatusFilter;
    return matchesQuery && matchesDept && matchesStatus;
  });
  const sortedTableRows = sortAttendanceRows(filteredTableRows, sort, data.monthlyLateCounts, warnCountMap);

  return (
    <section className="report-page">
      <section className="kpi-grid">
        <KpiCard
          icon={<UsersRound size={34} />}
          tone="green"
          label="พนักงานทั้งหมด"
          value={scopedTotal.toLocaleString()}
          unit="คน"
          note={selectedDeptLabel}
          progress={activeTotal ? Math.round((scopedCameToWork / activeTotal) * 100) : 0}
        />
        <KpiCard
          icon={<CheckCircle2 size={34} />}
          tone="blue"
          label="มาทำงาน"
          value={scopedPresent.toLocaleString()}
          unit="คน"
          note={`${presentPercent.toFixed(1)}% ของคนที่ต้องมา`}
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
          note={`${absentPercent.toFixed(1)}% ของคนที่ต้องมา`}
        />
        <KpiCard
          icon={<CalendarOff size={34} />}
          tone="gray"
          label="วันหยุด"
          value={scopedDayoff.toLocaleString()}
          unit="คน"
          note="ไม่นับในอัตราเข้างาน"
        />
        <KpiCard
          icon={<AlertTriangle size={34} />}
          tone="red"
          label="เสี่ยง"
          value={riskCount.toLocaleString()}
          unit="คน"
          note="สายสะสม ≥3 ครั้ง/เดือน"
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
            <span><i className="present" />มาตรงเวลา</span>
            <span><i className="late" />มาสาย</span>
            <span><i className="absent" />ขาดงาน</span>
          </div>
          <div className="stacked-bars">
            {data.deptRows.map((row, index) => {
              const deptActive = row.present + row.late + row.absent;
              const deptRate = deptActive > 0 ? Math.round(((row.present + row.late) / deptActive) * 100) : 0;
              const rateTone = deptRate >= 95 ? "good" : deptRate >= 85 ? "warn" : "bad";
              const pPresent = deptActive > 0 ? (row.present / deptActive) * 100 : 0;
              const pLate    = deptActive > 0 ? (row.late    / deptActive) * 100 : 0;
              const pAbsent  = deptActive > 0 ? (row.absent  / deptActive) * 100 : 0;
              return (
                <button
                  className={`stacked-row dept-click ${selectedDept === row.dept ? "active" : ""}`}
                  key={row.dept}
                  onClick={() => {
                    setSelectedDept(row.dept);
                    setDeptFilter("all");
                  }}
                  type="button"
                >
                  <span className="dept-rank">#{index + 1}</span>
                  <span className="stacked-dept-name">{row.dept}</span>
                  <div className="stacked-track">
                    <span className="seg present" style={{ width: `${pPresent}%` }} />
                    <span className="seg late"    style={{ width: `${pLate}%` }}>
                      {row.late > 0 ? `สาย ${row.late}` : ""}
                    </span>
                    <span className="seg absent"  style={{ width: `${pAbsent}%` }}>
                      {row.absent > 0 ? `ขาด ${row.absent}` : ""}
                    </span>
                  </div>
                  <span className={`dept-present-rate ${rateTone}`}>{deptRate}%</span>
                  <div className="stacked-row-end">
                    <strong className="dept-total">{deptActive} คน</strong>
                  </div>
                </button>
              );
            })}
            {data.deptRows.length === 0 ? <p className="empty-copy">ยังไม่มีข้อมูล report</p> : null}
          </div>
        </div>

        <div className="panel report-overview-card">
          <div className="sup-report-header">
            <h3>สถานะการตักเตือน{selectedDept !== "all" ? ` · ${selectedDept}` : ""}</h3>
            <span className={`sup-overall-badge ${supTotalLate === 0 || supTotalLate === supTotalWarned ? "done" : "pending"}`}>
              {supTotalLate === 0
                ? "ไม่มีคนมาสาย"
                : supTotalLate === supTotalWarned
                  ? "✓ ครบแล้ว"
                  : `⚠ ค้าง ${supTotalLate - supTotalWarned}`}
            </span>
          </div>
          <div className="sup-report-progress">
            <div className="sup-progress-bar">
              <div className="sup-progress-fill" style={{ width: `${supPct}%` }} />
            </div>
            <span className="sup-progress-pct">{supPct}%</span>
          </div>
          <div className="sup-report-stats">
            <div className="sup-report-stat-item">
              <span className="sup-report-stat-val" style={{ color: "#10b981" }}>{supTotalWarned}</span>
              <span className="sup-report-stat-lbl">เตือนแล้ว</span>
            </div>
            <div className="sup-report-stat-item">
              <span className="sup-report-stat-val" style={{ color: "#f59e0b" }}>{supTotalLate}</span>
              <span className="sup-report-stat-lbl">มาสายทั้งหมด</span>
            </div>
            <div className="sup-report-stat-item">
              <span className="sup-report-stat-val" style={{ color: supTotalLate - supTotalWarned > 0 ? "#ef4444" : "#94a3b8" }}>
                {supTotalLate - supTotalWarned}
              </span>
              <span className="sup-report-stat-lbl">ยังค้าง</span>
            </div>
          </div>
          <div className="sup-report-dept-list">
            {supDeptRows.map((row) => {
              const dPct = row.late ? Math.round((row.warned / row.late) * 100) : 100;
              const done = row.pending === 0;
              return (
                <div key={row.dept} className={`sup-dept-row ${done ? "done" : "pending"}`}>
                  <span className="sup-dept-name">{row.dept}</span>
                  <div className="sup-dept-bar-wrap">
                    <div className="sup-mini-bar">
                      <div className="sup-mini-fill" style={{ width: `${dPct}%` }} />
                    </div>
                    <span className="sup-mini-pct">{dPct}%</span>
                  </div>
                  <span className="sup-dept-counts">{row.warned}/{row.late}</span>
                  <span className={`sup-status-badge ${done ? "done" : "pending"}`}>
                    {done ? "✓" : `${row.pending}`}
                  </span>
                </div>
              );
            })}
            {supDeptRows.length === 0 && <p className="empty-copy">ไม่มีพนักงานมาสาย</p>}
          </div>
        </div>
      </section>

      <section className="panel report-table-panel">
        <div className="panel-collapse-trigger" onClick={() => setLateAbsentCollapsed(c => !c)}>
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <h3>รายละเอียด Late &amp; Absent</h3>
            {selectedDept !== "all" && <span className="dept-filter-chip">{selectedDept}</span>}
            <span className="table-count-badge">{sortedTableRows.length} คน</span>
          </div>
          <ChevronDown size={16} className="panel-collapse-chevron" style={{ transform: lateAbsentCollapsed ? "rotate(-90deg)" : "rotate(0deg)" }} />
        </div>
        {!lateAbsentCollapsed && (<>
        <div className="report-table-header" style={{ marginTop: "12px" }}>
          <div className="report-table-header-top">
            <div className="report-table-actions">
              <div className="status-tabs">
                {(["all", "Late", "Absent", "DayOff"] as const).map((s) => (
                  <button
                    key={s}
                    className={`status-tab${tableStatusFilter === s ? " active" : ""}${s === "Late" ? " amber" : s === "Absent" ? " red" : s === "DayOff" ? " slate" : ""}`}
                    onClick={() => setTableStatusFilter(s)}
                    type="button"
                  >
                    {s === "all" ? "ทั้งหมด" : s === "DayOff" ? "วันหยุด" : s}
                    <span className="tab-count">{s === "all" ? tableSourceRows.length : s === "Late" ? scopedLate : s === "Absent" ? scopedAbsent : scopedDayoff}</span>
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
                <th>เฉลี่ย/เดือน</th>
                <th>เสี่ยง</th>
                <th><SortButton columnKey="warnCount" setSort={setSort} sort={sort}>เตือนแล้ว</SortButton></th>
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
                    <td>
                      {row.status === "Absent" ? (
                        <select
                          className={`leave-select${leaveMap.has(row.empId) ? " saved" : ""}`}
                          value={leaveMap.get(row.empId) ?? ""}
                          onChange={(e) => saveLeave(row.empId, e.target.value)}
                        >
                          <option value="">— เลือกประเภท —</option>
                          <option value="ลาป่วย">ลาป่วย</option>
                          <option value="ลากิจ">ลากิจ</option>
                          <option value="ลาพักร้อน">ลาพักร้อน</option>
                          <option value="ขาดงาน">ขาดงาน</option>
                        </select>
                      ) : (
                        <span className={`status-pill ${row.status.toLowerCase()}`}>{STATUS_TH[row.status] ?? row.status}</span>
                      )}
                    </td>
                    <td>{row.status !== "Absent" && row.status !== "DayOff" ? <span className="late-minutes-badge">{formatLateTime(row.minutesLate)}</span> : "-"}</td>
                    <td>
                      <span className={monthly >= 3 ? "monthly-late-high" : ""}>
                        {monthly > 0 ? monthly : "-"}
                      </span>
                      <TrendBadge curr={monthly} prev={prevMonthLateCounts[row.empId] ?? 0} />
                    </td>
                    <td>
                      {monthly > 0 && monthlyLateMinutes[row.empId]
                        ? <span className="late-minutes-badge">{formatLateTime(Math.round(monthlyLateMinutes[row.empId] / monthly))}</span>
                        : "-"}
                    </td>
                    <td>{isRisk ? <span className="risk-badge">เสี่ยง</span> : null}</td>
                    <td>
                      {warnCount > 0 ? (
                        <details className="warn-history-details">
                          <summary className="warn-count-badge">✓ {warnCount} ครั้ง</summary>
                          <div className="warn-history-dates">
                            {(warnDates[row.empId] ?? []).map((d) => (
                              <span key={d} className="warn-date-chip">{formatDateTH(d)}</span>
                            ))}
                          </div>
                        </details>
                      ) : <span className="no-warn">-</span>}
                    </td>
                  </tr>
                );
              })}
              {sortedTableRows.length === 0 ? (
                <tr><td colSpan={12}>{reportData ? "ไม่มีข้อมูลตามเงื่อนไขที่เลือก" : "กด โหลดข้อมูล เพื่อสร้างรายงาน"}</td></tr>
              ) : null}
            </tbody>
          </table>
        </div>
        </>)}
      </section>

      {/* ── Supervisor Confirmation Summary ── */}
      {(() => {
        const totalDepts = (reportData?.deptRows ?? []).length;
        const confirmedCount = reportData ? deptConfirmations.size : 0;
        const pendingCount = totalDepts - confirmedCount;
        const pctConf = totalDepts ? Math.round((confirmedCount / totalDepts) * 100) : 0;
        const allConfirmed = totalDepts > 0 && pendingCount === 0;
        return (
          <section className="panel sup-check-panel">
            <div className="panel-collapse-trigger" onClick={() => setConfirmPanelCollapsed(c => !c)}>
              <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                <ClipboardCheck size={16} />
                <h3>สถานะการตรวจสอบการเข้างานรายหน่วยงาน</h3>
                {isoDate && <span className="table-count">{new Date(isoDate + "T00:00:00").toLocaleDateString("th-TH", { dateStyle: "medium" })}</span>}
              </div>
              <ChevronDown size={16} className="panel-collapse-chevron" style={{ transform: confirmPanelCollapsed ? "rotate(-90deg)" : "rotate(0deg)" }} />
            </div>
            {!confirmPanelCollapsed && (
              <>
                <div className="sup-check-header" style={{ marginTop: "12px" }}>
                  <div className="sup-check-title">
                    <p>ตรวจสอบว่าหัวหน้าหน่วยงานได้ยืนยันข้อมูลการเข้างานครบทุกหน่วยงานแล้วหรือยัง</p>
                  </div>
                  {reportData && (
                    <div className="sup-check-summary">
                      <div className="sup-check-stat">
                        <span className="sup-stat-value success">{confirmedCount}</span>
                        <span className="sup-stat-label">ยืนยันแล้ว</span>
                      </div>
                      <div className="sup-check-divider" />
                      <div className="sup-check-stat">
                        <span className="sup-stat-value muted">{totalDepts}</span>
                        <span className="sup-stat-label">ทั้งหมด</span>
                      </div>
                      <div className="sup-check-divider" />
                      <div className="sup-check-stat">
                        <span className={`sup-stat-value ${pendingCount > 0 ? "danger" : "success"}`}>{pendingCount}</span>
                        <span className="sup-stat-label">ยังค้าง</span>
                      </div>
                      <span className={`sup-overall-badge ${allConfirmed ? "done" : "pending"}`}>
                        {allConfirmed ? "✓ ครบแล้ว" : `⚠ ยังค้าง ${pendingCount} หน่วยงาน`}
                      </span>
                    </div>
                  )}
                </div>
                {reportData && (
                  <div className="sup-progress-wrap">
                    <div className="sup-progress-bar">
                      <div className="sup-progress-fill" style={{ width: `${pctConf}%` }} />
                    </div>
                    <span className="sup-progress-pct">{pctConf}%</span>
                  </div>
                )}
                {!reportData ? (
                  <p className="muted-text">กด โหลดข้อมูล เพื่อดูสถานะ</p>
                ) : (
                  <div className="sup-check-table-wrap">
                <table className="table compact-table">
                  <thead>
                    <tr>
                      <th>หน่วยงาน</th>
                      <th className="num">รวม</th>
                      <th className="num">ตรงเวลา</th>
                      <th className="num">มาสาย</th>
                      <th className="num">ขาด/ลา</th>
                      <th className="num">วันหยุด</th>
                      <th>สถานะ</th>
                      <th>ยืนยันโดย</th>
                      <th>เวลาที่ยืนยัน</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(reportData.deptRows ?? []).map(dept => {
                      const conf = deptConfirmations.get(dept.dept);
                      const done = !!conf;
                      return (
                        <tr key={dept.dept} className={done ? "sup-row-done" : "sup-row-pending"}>
                          <td><span className="dept-chip">{dept.dept}</span></td>
                          <td className="num">{dept.total}</td>
                          <td className="num">{dept.present}</td>
                          <td className="num">{dept.late > 0 ? <span className="csb late">{dept.late}</span> : "—"}</td>
                          <td className="num">{dept.absent > 0 ? <span className="csb absent">{dept.absent}</span> : "—"}</td>
                          <td className="num">{dept.dayoff > 0 ? dept.dayoff : "—"}</td>
                          <td>
                            {conf
                              ? <span className="confirm-status-badge confirmed"><CheckCircle2 size={12} />ยืนยันแล้ว</span>
                              : <span className="confirm-status-badge pending">รอยืนยัน</span>}
                          </td>
                          <td>{conf ? conf.confirmed_by : "—"}</td>
                          <td className="muted-text">{conf ? new Date(conf.confirmed_at).toLocaleString("th-TH", { dateStyle: "short", timeStyle: "short" }) : "—"}</td>
                        </tr>
                      );
                    })}
                    {(reportData.deptRows ?? []).length === 0 && (
                      <tr><td colSpan={9}>ไม่มีข้อมูลหน่วยงาน</td></tr>
                    )}
                  </tbody>
                  <tfoot>
                    <tr>
                      <td><strong>รวมทั้งหมด</strong></td>
                      <td className="num"><strong>{reportData.totalEmployees}</strong></td>
                      <td className="num"><strong>{reportData.present}</strong></td>
                      <td className="num"><strong>{reportData.late}</strong></td>
                      <td className="num"><strong>{reportData.absent}</strong></td>
                      <td className="num"><strong>{reportData.dayoff}</strong></td>
                      <td colSpan={3}>
                        <span className="confirm-summary-count">
                          {deptConfirmations.size}/{(reportData.deptRows ?? []).length} หน่วยงานยืนยันแล้ว
                        </span>
                      </td>
                    </tr>
                  </tfoot>
                </table>
                </div>
                )}
              </>
            )}
          </section>
        );
      })()}
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
  tone: "green" | "blue" | "amber" | "purple" | "gray";
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

function TrendBadge({ curr, prev }: { curr: number; prev: number }) {
  if (curr === 0 && prev === 0) return null;
  if (curr > prev) return <span className="trend-up" title={`+${curr - prev} จากเดือนก่อน`}>↑{curr - prev}</span>;
  if (curr < prev) return <span className="trend-down" title={`-${prev - curr} จากเดือนก่อน`}>↓{prev - curr}</span>;
  return <span className="trend-same" title="เท่ากับเดือนก่อน">→</span>;
}

function HeatmapPanel({ dailyStats, targetMonthKey }: { dailyStats: DailyStat[]; targetMonthKey: string }) {
  if (!targetMonthKey || dailyStats.length === 0) return null;
  const [yearNum, monthNum] = targetMonthKey.split("-").map(Number);
  const daysInMonth = new Date(yearNum, monthNum, 0).getDate();
  const byDate = Object.fromEntries(dailyStats.map((s) => [s.isoDate, s]));
  const thaiMonths = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];

  return (
    <div className="panel heatmap-panel">
      <div className="panel-title-row">
        <h3>สถิติรายวัน — {thaiMonths[monthNum]} {yearNum + 543}</h3>
        <div className="heatmap-legend">
          <span><i className="heatmap-dot good" />≥95%</span>
          <span><i className="heatmap-dot warn" />85–94%</span>
          <span><i className="heatmap-dot bad" />&lt;85%</span>
        </div>
      </div>
      <div className="heatmap-grid">
        {Array.from({ length: daysInMonth }, (_, i) => {
          const day = i + 1;
          const iso = `${yearNum}-${String(monthNum).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
          const stat = byDate[iso];
          const rate = stat && stat.total > 0 ? ((stat.present + stat.late) / stat.total) * 100 : null;
          const tone = rate === null ? "no-data" : rate >= 95 ? "good" : rate >= 85 ? "warn" : "bad";
          const tip = stat
            ? `${day}/${monthNum}: ${rate!.toFixed(0)}%  P:${stat.present} L:${stat.late} A:${stat.absent}`
            : `${day}/${monthNum}: ไม่มีข้อมูล`;
          return (
            <div key={iso} className={`heatmap-cell ${tone}`} title={tip}>
              <span className="heatmap-day">{day}</span>
              {rate !== null && <span className="heatmap-rate">{rate.toFixed(0)}%</span>}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function formatDateTH(isoDate: string) {
  const parts = isoDate.split("-");
  const m = Number(parts[1]);
  const d = Number(parts[2]);
  const thaiMonths = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  return `${d} ${thaiMonths[m]}`;
}
