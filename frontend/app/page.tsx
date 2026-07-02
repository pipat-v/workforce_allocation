"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import {
  AlertTriangle,
  ArrowRight,
  BarChart3,
  BookOpen,
  BriefcaseBusiness,
  CalendarClock,
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
  Search,
  Settings,
  Trash2,
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
  master_file_path: string | null;
  manpower_file_path: string | null;
  dayoff_shift_file_path: string | null;
};

type MasterFile = {
  id: string;
  file_type: MasterFileKey;
  file_path: string;
  original_filename: string | null;
  file_size_bytes: number | null;
  created_at: string;
  is_active: boolean;
};

type DayoffShiftEditorRow = {
  id: string;
  empId: string;
  name: string;
  dept: string;
  jobSite: string;
  dayoff: string;
  shift: string;
  shiftStart: string;
  shiftEnd: string;
  raw: Record<string, unknown>;
};

type ManpowerEditorRow = {
  id: string;
  dept: string;
  jobSite: string;
  shift: string;
  shiftStart: string;
  shiftEnd: string;
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
  jobSite: string;
  shift: string;
  shiftStart: string;
  dayoff: string;
  skill: string;
  level: number;
  origLevel: number;
};

type CombinedEmployeeRow = {
  empId: string;
  firstName: string;
  lastName: string;
  dept: string;
  position: string;
  jobSite: string;
  shift: string;
  shiftStart: string;
  shiftEnd: string;
  dayoff: string;
  skills: Record<string, number>;
};

type EmployeeDiff = {
  added: CombinedEmployeeRow[];
  removed: CombinedEmployeeRow[];
  changed: Array<{
    empId: string;
    name: string;
    dept: string;
    fields: Array<{ field: string; from: string; to: string }>;
  }>;
  unchangedCount: number;
  newRows: CombinedEmployeeRow[];
  newManpowerRows: Record<string, unknown>[];
};

type AttendanceRecord = {
  empId: string;
  name: string;
  dept: string;
  section: string;
  position: string;
  shift: string;
  shiftStart: string;
  shiftEnd: string;
  scanIn: string;      // HH:MM or "-"
  scanOut: string;     // HH:MM or "-"
  scanInDate: string;  // YYYY-MM-DD when different from isoTargetDate, else ""
  scanOutDate: string; // YYYY-MM-DD when different from isoTargetDate, else ""
  status: "Present" | "Late" | "Absent" | "DayOff" | "Pending" | "NoScanIn";
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
  dayoff: number;
  deptRows: Array<{ dept: string; present: number; late: number; absent: number; dayoff: number; total: number }>;
  lateRows: AttendanceRecord[];
  records: AttendanceRecord[];
  timestampRows: AttendanceRecord[];
  monthlyLateCounts: Record<string, number>;
  unmatchedScanIds: Array<{ empId: string; name: string; scanIn: string }>;
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
  | "ot"
  | "setting"
  | "help";

type MasterSubTab = "files" | "holidays" | "public_holidays" | "manpower" | "dayoff_shift";
type OtSubTab = "chart" | "summary" | "detail";

const navItems = [
  { id: "dashboard", label: "Dashboard", icon: HomeIcon },
  { id: "timestamp", label: "Upload Timestamp", icon: UploadCloud },
  { id: "results", label: "ผลลัพธ์การจัดสรร", icon: BriefcaseBusiness },
  { id: "timestamp_dept", label: "Timestamp With Dept", icon: Database },
  { id: "master", label: "Master Data", icon: FileSpreadsheet },
  { id: "skill", label: "Skill Matrix", icon: LayoutGrid },
  { id: "report", label: "Report & Dashboard", icon: BarChart3 },
  { id: "ot", label: "OT Dashboard", icon: TrendingUp },
  { id: "setting", label: "Setting", icon: Settings },
  { id: "help", label: "คู่มือการใช้งาน", icon: BookOpen },
];

type MasterFileKey = (typeof masterFileTypes)[number]["key"];
type MasterUploadState = Record<MasterFileKey, File | null>;
type SortDirection = "asc" | "desc";
type SortState = { key: string; direction: SortDirection } | null;

const publicWorkspace = "public";

export default function Home() {
  const [activeTab, setActiveTab] = useState<TabId>("dashboard");
  const [masterSubTab, setMasterSubTab] = useState<MasterSubTab>("files");
  const [otSubTab, setOtSubTab] = useState<OtSubTab>("chart");
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
  const [dashboardSectionFilter, setDashboardSectionFilter] = useState("all");
  const [detailStatusFilter, setDetailStatusFilter] = useState("all");
  const scrollToDetail = (filter: string) => {
    setDetailStatusFilter(filter);
    setTimeout(() => {
      document.getElementById("employee-detail-section")?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 50);
  };
  const [warnedIds, setWarnedIds] = useState<Set<string>>(new Set());
  const [warnPending, setWarnPending] = useState<Set<string>>(new Set());
  const [warnCountMap, setWarnCountMap] = useState<Record<string, number>>({});
  const [warnDates, setWarnDates] = useState<Record<string, string[]>>({});
  const [monthlyLateMinutes, setMonthlyLateMinutes] = useState<Record<string, number>>({});
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

  const latestRun = useMemo(
    () => (selectedRunId ? runs.find((r) => r.id === selectedRunId) : null) ?? runs.find((r) => !!r.scan_file_path),
    [selectedRunId, runs],
  );
  const reportSourceKey =
    activeMasterMap.employee_master?.file_path && latestRun?.scan_file_path
      ? [
          latestRun.id,
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

  const allSectionOptions = useMemo(() => {
    const base = dashboardDeptFilter === "all"
      ? (reportData?.records ?? [])
      : (reportData?.records ?? []).filter((r) => r.dept === dashboardDeptFilter);
    return Array.from(new Set(base.map((r) => r.section).filter(Boolean))).sort();
  }, [reportData, dashboardDeptFilter]);

  const dashboardReport = useMemo<ReportData | null>(() => {
    if (!reportData) return null;
    let filtered = reportData.records;
    if (dashboardDeptFilter !== "all") filtered = filtered.filter((r) => r.dept === dashboardDeptFilter);
    if (dashboardSectionFilter !== "all") filtered = filtered.filter((r) => r.section === dashboardSectionFilter);
    if (filtered === reportData.records) return reportData;
    const lateFiltered = filtered.filter((r) => r.status === "Late");
    const deptMap = new Map<string, { dept: string; present: number; late: number; absent: number; dayoff: number; total: number }>();
    for (const r of filtered) {
      const cur = deptMap.get(r.dept) ?? { dept: r.dept, present: 0, late: 0, absent: 0, dayoff: 0, total: 0 };
      cur.total += 1;
      if (r.status === "Present" || r.status === "NoScanIn") cur.present += 1;
      if (r.status === "Late") cur.late += 1;
      if (r.status === "Absent" || r.status === "Pending") cur.absent += 1;
      if (r.status === "DayOff") cur.dayoff += 1;
      deptMap.set(r.dept, cur);
    }
    return {
      ...reportData,
      records: filtered,
      timestampRows: filtered,
      totalEmployees: filtered.length,
      present: filtered.filter((r) => r.status === "Present" || r.status === "NoScanIn").length,
      late: lateFiltered.length,
      absent: filtered.filter((r) => r.status === "Absent" || r.status === "Pending").length,
      dayoff: filtered.filter((r) => r.status === "DayOff").length,
      lateRows: [...lateFiltered].sort((a, b) => b.minutesLate - a.minutesLate),
      deptRows: [...deptMap.values()],
    };
  }, [reportData, dashboardDeptFilter, dashboardSectionFilter]);

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
  const workDateKey = latestRun ? (latestRun.target_date ?? latestRun.created_at.slice(0, 10)) : null;
  const workDateBase = workDateKey ? new Date(workDateKey + "T00:00:00") : new Date();
  const workDate = workDateBase.toLocaleDateString("th-TH", {
    day: "numeric",
    month: "long",
    year: "numeric",
  });
  const [workTime, setWorkTime] = useState(() =>
    new Date().toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit", hour12: false }),
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
      setWorkTime(new Date().toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit", hour12: false }));
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
  }, [reportSourceKey, loadedReportKey]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => {
    if (!isoTargetDate) return;
    let cancelled = false;
    setWarnedIds(new Set());
    supabase
      .from("employee_warnings")
      .select("emp_id")
      .eq("warn_date", isoTargetDate)
      .then(({ data, error }) => {
        if (cancelled) return;
        if (error) { setError(error.message); return; }
        if (data) setWarnedIds(new Set(data.map((r: { emp_id: string }) => r.emp_id)));
      });
    return () => { cancelled = true; };
  }, [isoTargetDate]);

  useEffect(() => {
    let cancelled = false;
    void loadWarnCountMap(() => cancelled);
    return () => { cancelled = true; };
  }, [reportData?.targetMonthKey]);

  useEffect(() => {
    supabase
      .from("holidays")
      .select("date")
      .then(({ data, error }) => {
        if (error) { setError(error.message); return; }
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
      .select("id,file_type,file_path,original_filename,file_size_bytes,is_active,created_at")
      .order("created_at", { ascending: false });

    if (loadError) {
      setError(loadError.message);
      return;
    }

    const allFiles = (data ?? []) as MasterFile[];
    setMasterFileHistory(allFiles);

    const missingSize = allFiles.filter((f) => f.file_size_bytes == null);
    if (missingSize.length > 0) {
      const folderMap = new Map<string, MasterFile[]>();
      for (const f of missingSize) {
        const lastSlash = f.file_path.lastIndexOf("/");
        const folder = f.file_path.substring(0, lastSlash);
        if (!folderMap.has(folder)) folderMap.set(folder, []);
        folderMap.get(folder)!.push(f);
      }
      const updates: { id: string; file_size_bytes: number }[] = [];
      await Promise.all(
        Array.from(folderMap.entries()).map(async ([folder, files]) => {
          const { data: listed, error: listError } = await supabase.storage.from("workforce-inputs").list(folder);
          if (listError || !listed) return;
          for (const f of files) {
            const filename = f.file_path.substring(f.file_path.lastIndexOf("/") + 1);
            const found = listed.find((l) => l.name === filename);
            const size = found?.metadata?.size as number | undefined;
            if (size != null) updates.push({ id: f.id, file_size_bytes: size });
          }
        })
      );
      if (updates.length > 0) {
        await Promise.all(
          updates.map(({ id, file_size_bytes }) =>
            supabase.from("master_data_files").update({ file_size_bytes }).eq("id", id)
          )
        );
        setMasterFileHistory((prev) =>
          prev.map((f) => {
            const u = updates.find((u) => u.id === f.id);
            return u ? { ...f, file_size_bytes: u.file_size_bytes } : f;
          })
        );
      }
    }

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
      .select("id,target_date,status,scan_file_path,solver_status,original_filename,record_count,created_at,master_file_path,manpower_file_path,dayoff_shift_file_path")
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

    try {
      for (const item of masterFileTypes) {
        const file = masterUploads[item.key];
        if (!file) continue;

        const fileId = crypto.randomUUID();
        const path = `${publicWorkspace}/masters/${item.key}/${fileId}${getSafeFileExtension(file.name)}`;
        const { error: uploadError } = await supabase.storage
          .from("workforce-inputs")
          .upload(path, file, { upsert: true });

        if (uploadError) { setError(uploadError.message); return; }

        // Capture the current active record ID before deactivating for precise rollback
        const { data: prevActive } = await supabase
          .from("master_data_files")
          .select("id")
          .is("owner_id", null)
          .eq("file_type", item.key)
          .eq("is_active", true)
          .maybeSingle();
        const prevActiveId: string | null = prevActive?.id ?? null;

        const { error: deactivateError } = await supabase
          .from("master_data_files")
          .update({ is_active: false })
          .is("owner_id", null)
          .eq("file_type", item.key);

        if (deactivateError) { setError(deactivateError.message); return; }

        const { error: insertError } = await supabase
          .from("master_data_files")
          .insert({
            owner_id: null,
            file_type: item.key,
            file_path: path,
            original_filename: file.name,
            file_size_bytes: file.size,
            is_active: true,
          });

        if (insertError) {
          // Rollback: re-activate only the specific previous active record by ID
          if (prevActiveId) {
            await supabase
              .from("master_data_files")
              .update({ is_active: true })
              .eq("id", prevActiveId);
          }
          setError(insertError.message);
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
      await loadActiveMasters();
    } finally {
      setIsSavingMasters(false);
    }
  }

  async function saveDayoffShiftRows(rows: DayoffShiftEditorRow[]) {
    setError("");
    setMessage("");

    const fileId = crypto.randomUUID();
    const path = `${publicWorkspace}/masters/dayoff_shift/${fileId}.xlsx`;
    const workbookRows = rows.map((row) => {
      let raw: Record<string, unknown> = {
        ...row.raw,
        "User ID (Job Information)": row.empId,
        "ชื่อ นามสกุล": row.name,
      };
      raw = setRowCol(raw, row.dayoff, "วันหยุดประจำสัปดาห์", "วันหยุด", "dayoff", "Dayoff", "Day Off");
      raw = setRowCol(raw, row.shift, "อยู่กะไหน", "shift", "กะ", "Shift");
      raw = setRowCol(raw, row.shiftStart, "เวลาเข้างาน", "เวลาเข้า", "shift_start");
      raw = setRowCol(raw, row.shiftEnd, "เวลาออก", "เวลาออกงาน", "shift_end");
      raw = setRowCol(raw, row.jobSite, "หน่วยงานย่อย/Skill", "หน้างาน", "job_site", "Job Site");
      return raw;
    });
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
      throw new Error(uploadError.message);
    }

    const { data: prevActive } = await supabase
      .from("master_data_files")
      .select("id")
      .is("owner_id", null)
      .eq("file_type", "dayoff_shift")
      .eq("is_active", true)
      .maybeSingle();
    const prevActiveId: string | null = prevActive?.id ?? null;

    const { error: deactivateError } = await supabase
      .from("master_data_files")
      .update({ is_active: false })
      .is("owner_id", null)
      .eq("file_type", "dayoff_shift");

    if (deactivateError) {
      setError(deactivateError.message);
      throw new Error(deactivateError.message);
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
      if (prevActiveId) {
        await supabase.from("master_data_files").update({ is_active: true }).eq("id", prevActiveId);
      }
      setError(insertError.message);
      throw new Error(insertError.message);
    }

    setMessage("บันทึก Dayoff & Shift master แล้ว");
    await loadActiveMasters();
  }

  async function saveManpowerRows(rows: ManpowerEditorRow[]) {
    setError("");
    setMessage("");

    const fileId = crypto.randomUUID();
    const path = `${publicWorkspace}/masters/manpower_plan/${fileId}.xlsx`;
    const workbookRows = rows.map((row) => ({
      "หน่วยงาน": row.dept,
      "หน่วยงานย่อย": row.jobSite,
      "กะ": row.shift,
      "เวลาเข้า": row.shiftStart,
      "เวลาออก": row.shiftEnd,
    }));
    const worksheet = XLSX.utils.json_to_sheet(
      workbookRows.length ? workbookRows : [{ "หน่วยงาน": "", "หน่วยงานย่อย": "", "กะ": "", "เวลาเข้า": "", "เวลาออก": "" }],
    );
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Manpower Plan");
    const output = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const file = new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const { error: uploadError } = await supabase.storage
      .from("workforce-inputs")
      .upload(path, file, { upsert: true });

    if (uploadError) {
      setError(uploadError.message);
      throw new Error(uploadError.message);
    }

    const { data: prevActive } = await supabase
      .from("master_data_files")
      .select("id")
      .is("owner_id", null)
      .eq("file_type", "manpower_plan")
      .eq("is_active", true)
      .maybeSingle();
    const prevActiveId: string | null = prevActive?.id ?? null;

    const { error: deactivateError } = await supabase
      .from("master_data_files")
      .update({ is_active: false })
      .is("owner_id", null)
      .eq("file_type", "manpower_plan");

    if (deactivateError) {
      setError(deactivateError.message);
      throw new Error(deactivateError.message);
    }

    const { error: insertError } = await supabase
      .from("master_data_files")
      .insert({
        owner_id: null,
        file_type: "manpower_plan",
        file_path: path,
        original_filename: "manpower_plan-edited.xlsx",
        file_size_bytes: file.size,
        is_active: true,
      });

    if (insertError) {
      if (prevActiveId) {
        await supabase.from("master_data_files").update({ is_active: true }).eq("id", prevActiveId);
      }
      setError(insertError.message);
      throw new Error(insertError.message);
    }

    setMessage("บันทึก Manpower Plan แล้ว");
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
      throw new Error(uploadError.message);
    }

    const { data: prevActiveSkill } = await supabase
      .from("master_data_files")
      .select("id")
      .is("owner_id", null)
      .eq("file_type", "skill_matrix")
      .eq("is_active", true)
      .maybeSingle();
    const prevActiveSkillId: string | null = prevActiveSkill?.id ?? null;

    const { error: deactivateError } = await supabase
      .from("master_data_files")
      .update({ is_active: false })
      .is("owner_id", null)
      .eq("file_type", "skill_matrix");

    if (deactivateError) {
      setError(deactivateError.message);
      throw new Error(deactivateError.message);
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
      if (prevActiveSkillId) {
        await supabase.from("master_data_files").update({ is_active: true }).eq("id", prevActiveSkillId);
      }
      setError(insertError.message);
      throw new Error(insertError.message);
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
    setError("");
    const { error: deleteError } = await supabase
      .from("allocation_runs")
      .delete()
      .eq("id", run.id);
    if (deleteError) {
      setError(deleteError.message);
      return;
    }
    if (run.scan_file_path) {
      await supabase.storage.from("workforce-inputs").remove([run.scan_file_path]);
    }
    await loadRuns();
  }

  async function deleteMasterFile(file: MasterFile) {
    const label = file.original_filename ?? file.file_type;
    const activeWarning = file.is_active ? "\n⚠️ ไฟล์นี้กำลัง Active อยู่ — หลังลบระบบจะไม่มีข้อมูล Master สำหรับประเภทนี้" : "";
    if (!window.confirm(`ต้องการลบ "${label}" ใช่ไหม?\nการลบไม่สามารถย้อนกลับได้${activeWarning}`)) return;
    setError("");
    // Delete DB record first — if storage fails the record is gone (clean UI) but file is orphaned (harmless)
    // Reverse order risks: storage gone but DB record remains → user sees entry that 404s on download
    const { error: deleteError } = await supabase
      .from("master_data_files")
      .delete()
      .eq("id", file.id);
    if (deleteError) {
      setError(deleteError.message);
      return;
    }
    await supabase.storage.from("workforce-inputs").remove([file.file_path]);
    await loadActiveMasters();
  }

  async function createDailyRun() {
    setError("");
    setMessage("");
    setIsCreatingRun(true);

    try {
      if (!timestampFile) { setError("กรุณาเลือกไฟล์ timestamp"); return; }
      if (!hasAllActiveMasters) { setError("กรุณา upload master files ให้ครบก่อนสร้าง daily run"); return; }

      let recordCount: number | null = null;
      try {
        const buffer = await timestampFile.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        recordCount = XLSX.utils.sheet_to_json(firstSheet).length;
      } catch { /* ignore count error */ }

      const runId = crypto.randomUUID();
      const scanPath = `${publicWorkspace}/runs/${runId}/timestamp${getSafeFileExtension(timestampFile.name)}`;
      const { error: uploadError } = await supabase.storage
        .from("workforce-inputs")
        .upload(scanPath, timestampFile, { upsert: true });
      if (uploadError) { setError(uploadError.message); return; }

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
      if (insertError) { setError(insertError.message); return; }

      setTimestampFile(null);
      setMessage("สร้าง daily run แล้ว รอ worker ประมวลผล");
      await loadRuns();
    } catch (e) {
      setError(e instanceof Error ? e.message : "เกิดข้อผิดพลาด");
    } finally {
      setIsCreatingRun(false);
    }
  }

  async function loadReportDashboard() {
    setError("");
    setMessage("");
    setIsLoadingReport(true);

    try {
      const employeeMaster = activeMasterMap.employee_master;
      const latestRun = (selectedRunId ? runs.find(r => r.id === selectedRunId) : null) ?? runs.find(r => !!r.scan_file_path);

      if (!employeeMaster || !latestRun?.scan_file_path) {
        setError("ต้องมีไฟล์รายชื่อพนักงานและ timestamp ล่าสุดก่อนสร้าง Report & Dashboard");
        setIsLoadingReport(false);
        return;
      }

      // Use the master-file snapshot stored with this run (historical accuracy).
      // Fall back to the current active master only when the run pre-dates snapshot storage.
      const empPath      = latestRun.master_file_path        ?? employeeMaster.file_path;
      const manpowerPath = latestRun.manpower_file_path      ?? activeMasterMap.manpower_plan?.file_path;
      const dayoffPath   = latestRun.dayoff_shift_file_path  ?? activeMasterMap.dayoff_shift?.file_path;

      const [employeeRows, scanRows, manpowerRows, dayoffShiftRows] = await Promise.all([
        downloadSheetRows(empPath),
        downloadSheetRows(latestRun.scan_file_path),
        manpowerPath ? downloadSheetRows(manpowerPath) : Promise.resolve([]),
        dayoffPath   ? downloadSheetRows(dayoffPath)   : Promise.resolve([]),
      ]);

      // Find adjacent runs to supply night-shift clock-in/out times
      const currentDateKey = latestRun.target_date ?? latestRun.created_at.slice(0, 10);
      const prevRun = [...runs]
        .filter(
          (r) =>
            r.id !== latestRun.id &&
            !!r.scan_file_path &&
            (r.target_date ?? r.created_at.slice(0, 10)) < currentDateKey,
        )
        .sort((a, b) =>
          (b.target_date ?? b.created_at.slice(0, 10)).localeCompare(
            a.target_date ?? a.created_at.slice(0, 10),
          ),
        )[0];
      const nextRun = [...runs]
        .filter(
          (r) =>
            r.id !== latestRun.id &&
            !!r.scan_file_path &&
            (r.target_date ?? r.created_at.slice(0, 10)) > currentDateKey,
        )
        .sort((a, b) =>
          (a.target_date ?? a.created_at.slice(0, 10)).localeCompare(
            b.target_date ?? b.created_at.slice(0, 10),
          ),
        )[0];
      const [prevScanRows, nextScanRows] = await Promise.all([
        prevRun?.scan_file_path
          ? downloadSheetRows(prevRun.scan_file_path).catch(() => [])
          : Promise.resolve([] as Record<string, unknown>[]),
        nextRun?.scan_file_path
          ? downloadSheetRows(nextRun.scan_file_path).catch(() => [])
          : Promise.resolve([] as Record<string, unknown>[]),
      ]);

      const latestReport = buildReportData(employeeRows, scanRows, manpowerRows, dayoffShiftRows, holidayDates, prevScanRows, nextScanRows, currentDateKey);
      try {
        await saveTimestampWithDeptRows(latestRun.id, latestReport);
      } catch (saveError) {
        const errMsg =
          saveError instanceof Error
            ? saveError.message
            : (saveError as { message?: string })?.message ?? JSON.stringify(saveError);
        setError(`โหลด report ได้ แต่บันทึก Timestamp With Dept ไม่สำเร็จ: ${errMsg}`);
      }
      // Deduplicate runs by date key (keep latest upload per date) then sort ascending,
      // so that monthlyScanRows[mi-1] and [mi+1] are guaranteed to be adjacent calendar days.
      const runsByDate = new Map<string, typeof runs[0]>();
      for (const r of runs.filter((r) => !!r.scan_file_path)) {
        const dk = r.target_date ?? r.created_at.slice(0, 10);
        const existing = runsByDate.get(dk);
        if (!existing || r.created_at > existing.created_at) runsByDate.set(dk, r);
      }
      const sortedMonthlyRuns = [...runsByDate.values()].sort((a, b) =>
        (a.target_date ?? a.created_at.slice(0, 10)).localeCompare(
          b.target_date ?? b.created_at.slice(0, 10),
        ),
      );
      const monthlyLateCounts: Record<string, number> = {};
      const lateMinutesAcc: Record<string, number> = {};
      const prevMonthLateAcc: Record<string, number> = {};
      const [tY, tM] = latestReport.targetMonthKey.split("-").map(Number);
      const prevMonthKey = tM === 1
        ? `${tY - 1}-12`
        : `${tY}-${String(tM - 1).padStart(2, "0")}`;

      // Pre-fetch all unique master snapshots used across the monthly runs.
      // Runs that pre-date snapshot storage will have null paths and fall back to
      // the already-loaded employeeRows/manpowerRows/dayoffShiftRows from above.
      const masterRowsCache = new Map<string, Record<string, unknown>[]>();
      const uniquePaths = new Set(
        sortedMonthlyRuns.flatMap((r) =>
          [r.master_file_path, r.manpower_file_path, r.dayoff_shift_file_path].filter(Boolean) as string[]
        ),
      );
      await Promise.all(
        [...uniquePaths].map(async (p) => {
          try { masterRowsCache.set(p, await downloadSheetRows(p)); } catch { masterRowsCache.set(p, []); }
        }),
      );

      const monthlyScanRows = await Promise.all(
        sortedMonthlyRuns.map(async (run) => {
          try {
            return await downloadSheetRows(run.scan_file_path!);
          } catch {
            return [];
          }
        }),
      );

      for (let mi = 0; mi < monthlyScanRows.length; mi++) {
        const rows = monthlyScanRows[mi];
        const prevRows = monthlyScanRows[mi - 1] ?? [];
        const nextRows = monthlyScanRows[mi + 1] ?? [];
        const run = sortedMonthlyRuns[mi];
        const runEmpRows      = (run.master_file_path       ? masterRowsCache.get(run.master_file_path)       : null) ?? employeeRows;
        const runManpowerRows = (run.manpower_file_path     ? masterRowsCache.get(run.manpower_file_path)     : null) ?? manpowerRows;
        const runDayoffRows   = (run.dayoff_shift_file_path ? masterRowsCache.get(run.dayoff_shift_file_path) : null) ?? dayoffShiftRows;
        const dayReport = buildReportData(runEmpRows, rows, runManpowerRows, runDayoffRows, holidayDates, prevRows, nextRows, run.target_date ?? run.created_at.slice(0, 10));

        if (dayReport.targetMonthKey === latestReport.targetMonthKey) {
          for (const lateRow of dayReport.lateRows) {
            monthlyLateCounts[lateRow.empId] = (monthlyLateCounts[lateRow.empId] ?? 0) + 1;
            lateMinutesAcc[lateRow.empId] = (lateMinutesAcc[lateRow.empId] ?? 0) + lateRow.minutesLate;
          }
        } else if (dayReport.targetMonthKey === prevMonthKey) {
          for (const lateRow of dayReport.lateRows) {
            prevMonthLateAcc[lateRow.empId] = (prevMonthLateAcc[lateRow.empId] ?? 0) + 1;
          }
        }
      }

      setMonthlyLateMinutes(lateMinutesAcc);
      setPrevMonthLateCounts(prevMonthLateAcc);
      setReportData({
        ...latestReport,
        monthlyLateCounts,
      });
      setLoadedReportKey([
        latestRun.id,
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

  async function loadWarnCountMap(isCancelled?: () => boolean) {
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
    if (isCancelled?.()) return;
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
        target_date: report.isoTargetDate,
        emp_id: row.empId,
        name: row.name,
        dept: row.dept,
        position: row.position,
        shift: row.shift,
        shift_start: row.shiftStart,
        scan_in: row.scanIn,
        scan_out: row.scanOut,
        attendance_status: row.status,
        minutes_late: row.minutesLate,
      });
    }
    const rows = Array.from(rowMap.values());

    const batchErrors: string[] = [];
    for (let index = 0; index < rows.length; index += 500) {
      const { error: insertError } = await supabase
        .from("timestamp_with_dept")
        .upsert(rows.slice(index, index + 500), { onConflict: "run_id,emp_id" });
      if (insertError) batchErrors.push(insertError.message);
    }
    if (batchErrors.length > 0) throw new Error(batchErrors.join("; "));
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
            const isActive = activeTab === item.id;
            return (
              <div key={item.label}>
                <button
                  className={`nav-item ${isActive ? "active" : ""}`}
                  onClick={() => { setActiveTab(item.id as TabId); setError(""); setMessage(""); }}
                  type="button"
                >
                  <Icon size={19} />
                  <span>{item.label}</span>
                  {item.label === "Master Data" || item.label === "Report & Dashboard" || item.label === "OT Dashboard" ? (
                    <ChevronDown className={`nav-chevron${isActive ? " open" : ""}`} size={15} />
                  ) : null}
                </button>
                {item.id === "master" && isActive ? (
                  <div className="nav-sub-list">
                    {([
                      { id: "files", icon: FileSpreadsheet, label: "Master Files" },
                      { id: "manpower", icon: BarChart3, label: "Manpower" },
                      { id: "holidays", icon: CalendarDays, label: "วันพระ" },
                      { id: "public_holidays", icon: CalendarDays, label: "วันหยุดประจำปี" },
                      { id: "dayoff_shift", icon: CalendarClock, label: "Shift & Dayoff" },
                    ] as const).map((sub) => {
                      const SubIcon = sub.icon;
                      return (
                        <button
                          key={sub.id}
                          className={`nav-sub-item${masterSubTab === sub.id ? " active" : ""}${sub.id === "dayoff_shift" ? " primary" : ""}`}
                          onClick={() => setMasterSubTab(sub.id)}
                          type="button"
                        >
                          <SubIcon size={14} />
                          <span>{sub.label}</span>
                        </button>
                      );
                    })}
                  </div>
                ) : item.id === "ot" && isActive ? (
                  <div className="nav-sub-list">
                    {([
                      { id: "chart", icon: BarChart3, label: "แผนภูมิ" },
                      { id: "summary", icon: LayoutGrid, label: "สรุปรายหน่วยงาน" },
                      { id: "detail", icon: UsersRound, label: "สรุปรายพนักงาน" },
                    ] as const).map((sub) => {
                      const SubIcon = sub.icon;
                      return (
                        <button
                          key={sub.id}
                          className={`nav-sub-item${otSubTab === sub.id ? " active" : ""}`}
                          onClick={() => setOtSubTab(sub.id)}
                          type="button"
                        >
                          <SubIcon size={14} />
                          <span>{sub.label}</span>
                        </button>
                      );
                    })}
                  </div>
                ) : null}
              </div>
            );
          })}
        </nav>

        <button
          className="logout"
          type="button"
          onClick={() => { if (window.confirm("ต้องการออกจากระบบ?")) window.location.reload(); }}
        >
          <LogOut size={19} />
          <span>ออกจากระบบ</span>
        </button>
      </aside>

      <section className="main" data-tab={activeTab}>
        <header className="topbar">
          <div>
            <h1 className="title">Workforce Allocation System</h1>
            <p className="subtitle">ระบบจัดสรรตำแหน่งงานอัตโนมัติ</p>
          </div>
          <div className="top-actions">
            <div className="topbar-date-group">
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
              <div className="today-chip">
                <Clock size={14} />
                <div>
                  <span className="today-chip-label">วันนี้</span>
                  <strong>{todayDate}</strong>
                  <span className="work-time">{workTime} น.</span>
                </div>
              </div>
            </div>
          </div>
        </header>

        <section className="dashboard-head">
          <div className="dashboard-head-left">
            <h2>{activeNav?.label ?? "Dashboard"}</h2>
            {activeTab === "master" ? (
              <div className="master-sub-tabs">
                <button className={`master-sub-tab${masterSubTab === "files" ? " active" : ""}`} onClick={() => setMasterSubTab("files")}><FileSpreadsheet size={15} />Master Files</button>
                <button className={`master-sub-tab${masterSubTab === "manpower" ? " active" : ""}`} onClick={() => setMasterSubTab("manpower")}><BarChart3 size={15} />Manpower</button>
                <button className={`master-sub-tab${masterSubTab === "holidays" ? " active" : ""}`} onClick={() => setMasterSubTab("holidays")}><CalendarDays size={15} />วันพระ</button>
                <button className={`master-sub-tab${masterSubTab === "public_holidays" ? " active" : ""}`} onClick={() => setMasterSubTab("public_holidays")}><CalendarDays size={15} />วันหยุดประจำปี</button>
                <button className={`master-sub-tab master-sub-tab-primary${masterSubTab === "dayoff_shift" ? " active" : ""}`} onClick={() => setMasterSubTab("dayoff_shift")}><CalendarClock size={15} />Shift & Dayoff</button>
              </div>
            ) : activeTab === "ot" ? (
              <div className="master-sub-tabs">
                <button className={`master-sub-tab${otSubTab === "chart" ? " active" : ""}`} onClick={() => setOtSubTab("chart")}><BarChart3 size={15} />แผนภูมิ</button>
                <button className={`master-sub-tab${otSubTab === "summary" ? " active" : ""}`} onClick={() => setOtSubTab("summary")}><LayoutGrid size={15} />สรุปรายหน่วยงาน</button>
                <button className={`master-sub-tab${otSubTab === "detail" ? " active" : ""}`} onClick={() => setOtSubTab("detail")}><UsersRound size={15} />สรุปรายพนักงาน</button>
              </div>
            ) : null}
            {activeTab === "dashboard" && allDeptOptions.length > 0 ? (
              <select
                aria-label="กรองหน่วยงาน"
                className="dept-filter-select"
                value={dashboardDeptFilter}
                onChange={(e) => { setDashboardDeptFilter(e.target.value); setDashboardSectionFilter("all"); }}
              >
                <option value="all">ทุกหน่วยงาน</option>
                {allDeptOptions.map((dept) => (
                  <option key={dept} value={dept}>{dept}</option>
                ))}
              </select>
            ) : null}
            {activeTab === "dashboard" && allSectionOptions.length > 0 ? (
              <select
                aria-label="กรองหน่วยงานย่อย"
                className="dept-filter-select"
                value={dashboardSectionFilter}
                onChange={(e) => setDashboardSectionFilter(e.target.value)}
              >
                <option value="all">ทุกหน่วยงานย่อย</option>
                {allSectionOptions.map((s) => (
                  <option key={s} value={s}>{s}</option>
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
                onClick={() => scrollToDetail("all")}
              />
              <KpiCard
                icon={<ClipboardCheck size={34} />}
                tone="blue"
                label="ตรงเวลา"
                value={presentPeople.toLocaleString()}
                unit="คน"
                note={`${pctPresent}% ของพนักงานทั้งหมด`}
                onClick={() => scrollToDetail("Present")}
              />
              <KpiCard
                icon={<Clock size={34} />}
                tone="amber"
                label="มาสาย"
                value={latePeople.toLocaleString()}
                unit="คน"
                note={`${pctLate}% ของพนักงานทั้งหมด`}
                onClick={() => document.getElementById("late-people-section")?.scrollIntoView({ behavior: "smooth", block: "start" })}
              />
              <KpiCard
                icon={<UserX size={34} />}
                tone="purple"
                label="ขาด/ลา"
                value={absentPeople.toLocaleString()}
                unit="คน"
                note={`${pctAbsent}% ของพนักงานทั้งหมด`}
                onClick={() => scrollToDetail("Absent")}
              />
              <KpiCard
                icon={<CalendarOff size={34} />}
                tone="gray"
                label="วันหยุด"
                value={dayoffPeople.toLocaleString()}
                unit="คน"
                note={`${pctDayoff}% ของพนักงานทั้งหมด`}
                onClick={() => scrollToDetail("DayOff")}
              />
            </section>

            <DashboardPanels
              activeMasterMap={activeMasterMap}
              assignedPeople={presentPeople}
              dashboardDeptFilter={dashboardDeptFilter}
              dashboardSectionFilter={dashboardSectionFilter}
              detailStatusFilter={detailStatusFilter}
              setDetailStatusFilter={setDetailStatusFilter}
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
            isSavingMasters={isSavingMasters}
            masterFileHistory={masterFileHistory}
            masterSubTab={masterSubTab}
            masterUploads={masterUploads}
            onDeleteMasterFile={deleteMasterFile}
            onHolidaysChanged={(dates) => setHolidayDates(dates)}
            onMastersSaved={loadActiveMasters}
            saveDayoffShiftRows={saveDayoffShiftRows}
            saveManpowerRows={saveManpowerRows}
            saveMasterFiles={saveMasterFiles}
            setError={setError}
            setMessage={setMessage}
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
            isLoadingReport={isLoadingReport}
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
            dayoffShiftFile={activeMasterMap.dayoff_shift}
            employeeMasterFile={activeMasterMap.employee_master}
            saveSkillMatrixRows={saveSkillMatrixRows}
          />
        ) : null}

        {activeTab === "ot" ? (
          <OTDashboard
            reportData={reportData}
            activeMasterMap={activeMasterMap}
            isLoadingReport={isLoadingReport}
            otSubTab={otSubTab}
            setOtSubTab={setOtSubTab}
            scanUploadedAt={latestRun?.created_at ?? null}
            holidayDates={holidayDates}
          />
        ) : null}

        {activeTab === "help" ? (
          <HelpGuidePage setActiveTab={setActiveTab} setMasterSubTab={setMasterSubTab} setOtSubTab={setOtSubTab} />
        ) : null}

        {!["dashboard", "master", "timestamp", "results", "timestamp_dept", "report", "skill", "ot", "help"].includes(activeTab) ? (
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
  if (error || !data?.signedUrl) {
    alert(`ดาวน์โหลดไม่สำเร็จ: ${error?.message ?? "ไม่พบ URL"}`);
    return;
  }
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
  // raw:true stops SheetJS from auto-guessing ambiguous CSV date strings (e.g. "02-07-2026" as
  // US-style MM-DD when day <= 12, silently producing the wrong month). Real .xlsx date cells are
  // unaffected (their type is explicit in the file); text stays text for parseTimestamp to parse
  // as DD-MM-YYYY, and real date cells still come through via cellDates.
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true, raw: true });
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
  prevDayScanRows: Record<string, unknown>[] = [],
  nextDayScanRows: Record<string, unknown>[] = [],
  explicitTargetDate?: string,
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
  const manpowerLookup = buildManpowerLookup(parseManpowerRows(manpowerRows));
  const scanByEmp = new Map<string, { name: string; times: Date[]; seenMs: Set<number> }>();

  for (const row of scanRows) {
    const empId = cleanEmpId(row["Employee ID"] ?? row["Emp ID"] ?? row["รหัสพนักงาน"]);
    const timestamp = parseTimestamp(row["Timestamp"]);
    if (!empId || !timestamp) continue;

    const current = scanByEmp.get(empId) ?? {
      name: String(row["Employee Name"] ?? row["name"] ?? "").trim(),
      times: [],
      seenMs: new Set<number>(),
    };
    const ms = timestamp.getTime();
    if (!current.seenMs.has(ms)) {
      current.seenMs.add(ms);
      current.times.push(timestamp);
    }
    scanByEmp.set(empId, current);
  }

  const latestTimestamp = Array.from(scanByEmp.values())
    .flatMap((entry) => entry.times)
    .sort((a, b) => b.getTime() - a.getTime())[0];

  // ใช้วันที่ของ run เป็นหลักถ้ามีระบุมาให้ (ตรงกับวันที่ผู้ใช้เลือก/อัปโหลด) — ไฟล์สแกนจริงมักมีข้อมูล
  // ของ "เมื่อวาน+วันนี้" ปนกัน (export ตอนเช้า) ทำให้ยอดสแกนของเมื่อวานมักเยอะกว่าวันนี้ (ที่เพิ่งผ่านไปไม่กี่ชม.)
  // ถ้าใช้ "โหมดของวันที่สแกน" (mode) จะได้เมื่อวานผิดวันเสมอ จึงต้องยึดวันที่ของ run เป็นหลักแทน
  // ยังคง fallback เป็น mode-of-scan-dates ไว้กรณีไม่มีวันที่ระบุมา (เช่น legacy call site)
  const allScanTimes = Array.from(scanByEmp.values()).flatMap((e) => e.times);
  const isoCounts = new Map<string, number>();
  for (const t of allScanTimes) {
    const iso = `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;
    isoCounts.set(iso, (isoCounts.get(iso) ?? 0) + 1);
  }
  const isoTargetDate = explicitTargetDate
    ? explicitTargetDate
    : isoCounts.size > 0
      ? [...isoCounts.entries()].sort((a, b) => b[1] - a[1])[0][0]
      : latestTimestamp
        ? `${latestTimestamp.getFullYear()}-${String(latestTimestamp.getMonth() + 1).padStart(2, "0")}-${String(latestTimestamp.getDate()).padStart(2, "0")}`
        : "";
  const targetTimestamp = isoTargetDate ? new Date(`${isoTargetDate}T12:00:00`) : latestTimestamp;
  const targetDate = targetTimestamp?.toLocaleDateString("th-TH") ?? "-";
  const targetMonthKey = isoTargetDate ? isoTargetDate.slice(0, 7) : "";
  const _now = new Date();
  const todayIso = `${_now.getFullYear()}-${String(_now.getMonth() + 1).padStart(2, "0")}-${String(_now.getDate()).padStart(2, "0")}`;
  const nowMinutes = _now.getHours() * 60 + _now.getMinutes();
  const isTargetToday = isoTargetDate === todayIso;

  // For night shift workers (e.g. 22:00–07:00) whose only scans today are early-morning
  // clock-outs, pull in the previous day's evening scans (≥20:00) as the real clock-in.
  // Only applied to employees whose entire current-day scan set is before 12:00, which
  // distinguishes them from day/afternoon workers who also have afternoon scans today.
  if (prevDayScanRows.length > 0) {
    const prevEveningByEmp = new Map<string, { times: Date[]; seenMs: Set<number> }>();
    for (const row of prevDayScanRows) {
      const empId = cleanEmpId(row["Employee ID"] ?? row["Emp ID"] ?? row["รหัสพนักงาน"]);
      const ts = parseTimestamp(row["Timestamp"]);
      if (!empId || !ts || ts.getHours() < 15) continue; // ≥15:00 covers 17:00 evening shifts
      const entry2 = prevEveningByEmp.get(empId) ?? { times: [], seenMs: new Set<number>() };
      const ms = ts.getTime();
      if (!entry2.seenMs.has(ms)) { entry2.seenMs.add(ms); entry2.times.push(ts); }
      prevEveningByEmp.set(empId, entry2);
    }
    for (const [empId, entry] of scanByEmp.entries()) {
      if (entry.times.length === 0 || !entry.times.every((t) => t.getHours() < 12)) continue;
      const prevTimes = prevEveningByEmp.get(empId)?.times ?? [];
      if (prevTimes.length > 0) entry.times = [...prevTimes, ...entry.times];
    }
  }

  // For night shift workers who clocked in this evening but have no early-morning
  // clock-out in today's file, pull the next day's early-morning scans as clock-out.
  // This lets the previous day's report show the correct scanOut even when the
  // machine scan file only captures up to midnight.
  if (nextDayScanRows.length > 0) {
    const nextMorningByEmp = new Map<string, { times: Date[]; seenMs: Set<number> }>();
    for (const row of nextDayScanRows) {
      const empId = cleanEmpId(row["Employee ID"] ?? row["Emp ID"] ?? row["รหัสพนักงาน"]);
      const ts = parseTimestamp(row["Timestamp"]);
      if (!empId || !ts || ts.getHours() >= 12) continue; // early-morning only
      const entry2 = nextMorningByEmp.get(empId) ?? { times: [], seenMs: new Set<number>() };
      const ms = ts.getTime();
      if (!entry2.seenMs.has(ms)) { entry2.seenMs.add(ms); entry2.times.push(ts); }
      nextMorningByEmp.set(empId, entry2);
    }
    for (const [empId, entry] of scanByEmp.entries()) {
      const hasEveningScan = entry.times.some((t) => t.getHours() >= 15); // ≥15:00 covers 17:00 evening shifts
      const hasMorningScan = entry.times.some((t) => t.getHours() < 12);
      if (!hasEveningScan || hasMorningScan) continue;
      const nextMorning = nextMorningByEmp.get(empId)?.times ?? [];
      if (nextMorning.length > 0) entry.times = [...entry.times, ...nextMorning];
    }
  }


  const baseEmployees = employees.length
    ? employees
    : Array.from(scanByEmp.entries()).map(([empId, entry]) => ({
        empId,
        name: entry.name || empId,
        dept: "ไม่ระบุ",
        position: "พนักงาน",
      }));

  const unmatchedScanIds = employees.length
    ? Array.from(scanByEmp.entries())
        .filter(([empId]) => !employeeMap.has(empId))
        .map(([empId, entry]) => {
          const earliest = entry.times.sort((a, b) => a.getTime() - b.getTime())[0];
          return {
            empId,
            name: entry.name || "-",
            scanIn: earliest ? toTimeText(earliest) : "-",
          };
        })
        .sort((a, b) => a.empId.localeCompare(b.empId))
    : [];

  const records: AttendanceRecord[] = baseEmployees.flatMap((employee): AttendanceRecord[] => {
    const dayoffShift = dayoffShiftMap.get(employee.empId);
    const scans = scanByEmp.get(employee.empId)?.times ?? [];
    const shift = normalizeShiftLabel(dayoffShift?.shift) || "กะ 1";
    const manpowerEntry = lookupManpowerTime(manpowerLookup, employee.dept, dayoffShift?.section ?? "", shift);
    const shiftStart = dayoffShift?.shiftStart || manpowerEntry?.shiftStart || "07:00";
    const shiftEnd = dayoffShift?.shiftEnd || manpowerEntry?.shiftEnd || "";
    // Use the known shift start time to pick the correct clock-in timestamp.
    // For night shifts (e.g. shiftStart=22:00) this selects the evening scan,
    // not the early-morning clock-out that a plain min() would return.
    const scanIn = findScanIn(scans, shiftStart, isoTargetDate);
    const isScheduledOff = targetTimestamp
      ? isEmployeeDayOff(dayoffShift?.dayoff, targetTimestamp, holidaySet)
      : false;
    if (isScheduledOff) return [{
      empId: employee.empId,
      name: employee.name,
      dept: employee.dept,
      section: dayoffShift?.section ?? "",
      position: employee.position,
      shift,
      shiftStart,
      shiftEnd,
      scanIn: "-",
      scanOut: "-",
      scanInDate: "",
      scanOutDate: "",
      status: "DayOff" as const,
      minutesLate: 0,
    }];

    const [ssH = 7, ssM = 0] = shiftStart.split(":").map(Number);
    const shiftStartMinutes = ssH * 60 + ssM;

    // Detect "forgot to scan in": no valid clock-in, but there is a scan on the
    // target date that falls outside the clock-in window (a probable clock-out).
    let scanOut = findScanOut(scans, scanIn);
    let noScanIn = false;
    if (!scanIn && isoTargetDate) {
      const toIsoStr = (t: Date) =>
        `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;
      const postWindow = scans.filter(
        (t) => toIsoStr(t) === isoTargetDate && !isInClockInWindow(t, ssH, ssM, 360)
      );
      if (postWindow.length > 0) {
        noScanIn = true;
        scanOut = postWindow.sort((a, b) => b.getTime() - a.getTime())[0];
      }
    }

    const minutesLate = scanIn ? Math.max(0, minutesBetween(shiftStart, scanIn)) : 0;
    // If we're in the early hours (before 06:00) and the shift starts in the evening
    // (18:00+), we've crossed midnight — the shift already started. Without this guard,
    // nowMinutes=30 < shiftStartMinutes=1320 would incorrectly show Pending at 00:30 AM
    // for a 22:00 shift that started 2.5 hours ago.
    const crossedMidnight = nowMinutes < 360 && shiftStartMinutes >= 1080;
    const shiftNotStarted = isTargetToday && !scanIn && !noScanIn && !crossedMidnight && nowMinutes < shiftStartMinutes;
    const status = !scanIn
      ? (noScanIn ? "NoScanIn" : shiftNotStarted ? "Pending" : "Absent")
      : minutesLate > 5 ? "Late" : "Present";

    const toIsoDate = (t: Date) =>
      `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;
    const scanInDateVal = scanIn ? toIsoDate(scanIn) : "";
    const scanOutDateVal = scanOut ? toIsoDate(scanOut) : "";
    return [{
      empId: employee.empId,
      name: employee.name,
      dept: employee.dept,
      section: dayoffShift?.section ?? "",
      position: employee.position,
      shift,
      shiftStart,
      shiftEnd,
      scanIn: scanIn ? toTimeText(scanIn) : "-",
      scanOut: scanOut ? toTimeText(scanOut) : "-",
      scanInDate: scanInDateVal !== isoTargetDate ? scanInDateVal : "",
      scanOutDate: scanOutDateVal !== isoTargetDate ? scanOutDateVal : "",
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
    if (record.status === "Absent" || record.status === "Pending") current.absent += 1;
    if (record.status === "NoScanIn") current.present += 1;
    if (record.status === "DayOff") current.dayoff += 1;
    deptMap.set(record.dept, current);
  }

  return {
    targetDate,
    isoTargetDate,
    targetMonthKey,
    totalEmployees: records.length,
    present: records.filter((record) => record.status === "Present" || record.status === "NoScanIn").length,
    late: records.filter((record) => record.status === "Late").length,
    absent: records.filter((record) => record.status === "Absent" || record.status === "Pending").length,
    dayoff: records.filter((record) => record.status === "DayOff").length,
    deptRows: Array.from(deptMap.values())
      .sort((a, b) => b.total - a.total),
    lateRows: records
      .filter((record) => record.status === "Late")
      .sort((a, b) => b.minutesLate - a.minutesLate),
    records,
    timestampRows: records,
    monthlyLateCounts: {},
    unmatchedScanIds,
  };
}

type ManpowerRow = { dept: string; jobSite: string; shift: string; shiftStart: string; shiftEnd: string };
type ManpowerLookupEntry = { shiftStart: string; shiftEnd: string };
type ManpowerLookup = {
  specific: Map<string, ManpowerLookupEntry>;
  byShift: Map<string, ManpowerLookupEntry>;
  byDept: Map<string, ManpowerLookupEntry>;
};

function parseManpowerRows(rows: Record<string, unknown>[]): ManpowerRow[] {
  return rows
    .map((row) => ({
      dept: String(row["หน่วยงาน"] ?? row["dept"] ?? "").trim(),
      jobSite: String(row["หน่วยงานย่อย"] ?? row["หน้างาน"] ?? row["job_site"] ?? "").trim(),
      shift: normalizeShiftLabel(row["กะ"] ?? row["shift"] ?? row["อยู่กะไหน"]),
      shiftStart: normalizeTimeText(row["เวลาเข้า"] ?? row["shift_start"] ?? row["เวลาเข้างาน"]),
      shiftEnd: normalizeTimeText(row["เวลาออก"] ?? row["เวลาออกงาน"] ?? row["shift_end"]),
    }))
    .filter((r) => r.dept && (r.shiftStart || r.shiftEnd));
}

// สามชั้น: เจาะจงหน่วยงาน+หน่วยงานย่อย+กะ -> เจาะจงหน่วยงาน+กะ (ไม่ระบุหน่วยงานย่อย) -> ค่า default ของทั้งหน่วยงาน
function buildManpowerLookup(rows: ManpowerRow[]): ManpowerLookup {
  const specific = new Map<string, ManpowerLookupEntry>();
  const byShift = new Map<string, ManpowerLookupEntry>();
  const byDept = new Map<string, ManpowerLookupEntry>();
  for (const r of rows) {
    if (!r.dept) continue;
    const entry: ManpowerLookupEntry = { shiftStart: r.shiftStart, shiftEnd: r.shiftEnd };

    const deptKey = makeDeptShiftKey(r.dept, "");
    if (!byDept.has(deptKey)) byDept.set(deptKey, entry);

    if (r.shift) {
      const shiftKey = makeDeptShiftKey(r.dept, r.shift);
      if (!byShift.has(shiftKey)) byShift.set(shiftKey, entry);

      if (r.jobSite) {
        const specificKey = `${r.dept}||${r.jobSite}||${normalizeShiftKey(r.shift)}`;
        if (!specific.has(specificKey)) specific.set(specificKey, entry);
      }
    }
  }
  return { specific, byShift, byDept };
}

function lookupManpowerTime(lookup: ManpowerLookup, dept: string, jobSite: string, shift: string): ManpowerLookupEntry | undefined {
  // "ผู้จัดการ" หมายถึงไม่มีกะตายตัว ไม่ใช่ชื่อกะจริง ต้องไม่ resolve เวลาให้ (เหมือนไม่มีกะเลย)
  if (!shift || shift === "ผู้จัดการ") return undefined;
  if (jobSite && shift) {
    const hit = lookup.specific.get(`${dept}||${jobSite}||${normalizeShiftKey(shift)}`);
    if (hit) return hit;
  }
  if (shift) {
    const hit = lookup.byShift.get(makeDeptShiftKey(dept, shift));
    if (hit) return hit;
  }
  return lookup.byDept.get(makeDeptShiftKey(dept, ""));
}

function findRowCol(row: Record<string, unknown>, ...targets: string[]): string {
  const val = findRowColRaw(row, ...targets);
  return val === undefined ? "" : String(val).trim();
}

// Like findRowCol but returns the original value (Date/number/string) instead
// of stringifying it, so callers like normalizeTimeText can still tell a
// Date or Excel serial number apart from plain text.
function findRowColRaw(row: Record<string, unknown>, ...targets: string[]): unknown {
  const norm = (s: string) => s.replace(/[\s\r\n]+/g, "").toLowerCase();
  const normedTargets = targets.map(norm);
  for (const [key, val] of Object.entries(row)) {
    const k = norm(key);
    if (normedTargets.some((t) => k === t)) return val;
  }
  return undefined;
}

// Writes `value` into whichever existing key in `row` matches one of `targets`
// (case/whitespace-insensitive), so edits land on the same column findRowCol
// reads from instead of creating a duplicate column under a new key name.
// Also drops any other keys matching the same targets, since some master
// files already carry duplicate columns (e.g. a header with \r\n vs \n)
// from earlier saves — without this, findRowCol could keep reading the
// other stale duplicate instead of the one just written.
function setRowCol(row: Record<string, unknown>, value: string, ...targets: string[]): Record<string, unknown> {
  const norm = (s: string) => s.replace(/[\s\r\n]+/g, "").toLowerCase();
  const normedTargets = targets.map(norm);
  const next: Record<string, unknown> = {};
  let written = false;
  for (const [key, val] of Object.entries(row)) {
    if (normedTargets.some((t) => norm(key) === t)) {
      if (!written) {
        next[key] = value;
        written = true;
      }
      continue;
    }
    next[key] = val;
  }
  if (!written) next[targets[0]] = value;
  return next;
}

function buildDayoffShiftMap(rows: Record<string, unknown>[]) {
  const map = new Map<string, { dayoff: string; shift: string; shiftStart: string; shiftEnd: string; section: string }>();
  for (const row of rows) {
    const empId = cleanEmpId(
      row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"]
    );
    if (!empId) continue;
    map.set(empId, {
      dayoff: findRowCol(row, "วันหยุดประจำสัปดาห์", "วันหยุด", "dayoff", "Dayoff", "Day Off"),
      shift: findRowCol(row, "อยู่กะไหน", "shift", "กะ", "Shift"),
      shiftStart: normalizeTimeText(findRowCol(row, "เวลาเข้างาน", "เวลาเข้า", "shift_start")),
      shiftEnd: normalizeTimeText(findRowCol(row, "เวลาออก", "เวลาออกงาน", "shift_end")),
      section: findRowCol(row, "หน่วยงานย่อย/Skill", "หน้างาน", "job_site", "Job Site"),
    });
  }
  return map;
}

function toDayoffShiftEditorRow(row: Record<string, unknown>, index: number): DayoffShiftEditorRow {
  const empId = cleanEmpId(row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"]);
  const firstName = String(row["First Name (Local)"] ?? "").trim();
  const lastName = String(row["Last Name (Local)"] ?? "").trim();
  const fallbackName = String(row["ชื่อ นามสกุล"] ?? row["Employee Name"] ?? row["Name"] ?? "").trim();
  const shiftStartRaw = findRowColRaw(row, "เวลาเข้างาน", "เวลาเข้า", "shift_start");
  const shiftStart = normalizeTimeText(shiftStartRaw);
  const shiftEnd = normalizeTimeText(findRowCol(row, "เวลาออก", "เวลาออกงาน", "shift_end")) || (shiftStart ? addHoursToTime(shiftStart, 9) : "");
  return {
    id: `${empId || "row"}-${index}`,
    empId,
    name: `${firstName} ${lastName}`.trim() || fallbackName || empId,
    dept: findRowCol(row, "หน่วยงาน", "Org. Unit Description", "Name (Section)", "แผนก", "Department"),
    jobSite: findRowCol(row, "หน่วยงานย่อย/Skill", "หน้างาน", "job_site", "Job Site"),
    dayoff: findRowCol(row, "วันหยุดประจำสัปดาห์", "วันหยุด", "dayoff", "Dayoff", "Day Off"),
    shift: findRowCol(row, "อยู่กะไหน", "shift", "กะ", "Shift"),
    shiftStart,
    shiftEnd,
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

function TimeInput24({
  value,
  onChange,
  className,
  style,
  title,
}: {
  value: string;
  onChange: (v: string) => void;
  className?: string;
  style?: React.CSSProperties;
  title?: string;
}) {
  const [open, setOpen] = useState(false);
  const [dropPos, setDropPos] = useState({ top: 0, left: 0 });
  const parsed = value.match(/^(\d{1,2}):(\d{2})$/);
  const selH = parsed ? Number(parsed[1]) : 0;
  const selM = parsed ? Number(parsed[2]) : 0;
  const btnRef = useRef<HTMLButtonElement>(null);
  const dropRef = useRef<HTMLDivElement>(null);
  const hourRef = useRef<HTMLDivElement>(null);
  const minRef = useRef<HTMLDivElement>(null);

  const calcPos = () => {
    if (btnRef.current) {
      const rect = btnRef.current.getBoundingClientRect();
      setDropPos({ top: rect.bottom + 4, left: rect.left });
    }
  };

  useEffect(() => {
    if (!open) return;
    let rafId = requestAnimationFrame(() => {
      const hEl = hourRef.current?.children[selH] as HTMLElement | undefined;
      const mEl = minRef.current?.children[selM] as HTMLElement | undefined;
      hEl?.scrollIntoView({ block: "center" });
      mEl?.scrollIntoView({ block: "center" });
    });
    const handleDown = (e: MouseEvent) => {
      const clickedPortal = dropRef.current?.contains(e.target as Node);
      const clickedBtn = btnRef.current?.contains(e.target as Node);
      if (!clickedPortal && !clickedBtn) setOpen(false);
    };
    const handleScroll = () => calcPos();
    document.addEventListener("mousedown", handleDown);
    window.addEventListener("scroll", handleScroll, true);
    window.addEventListener("resize", handleScroll);
    return () => {
      cancelAnimationFrame(rafId);
      document.removeEventListener("mousedown", handleDown);
      window.removeEventListener("scroll", handleScroll, true);
      window.removeEventListener("resize", handleScroll);
    };
  }, [open, selH, selM]);

  const handleToggle = () => {
    if (!open) calcPos();
    setOpen((o) => !o);
  };

  const pick = (h: number, m: number, close = false) => {
    onChange(`${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`);
    if (close) setOpen(false);
  };

  const dropdown = open && typeof document !== "undefined"
    ? createPortal(
        <div
          ref={dropRef}
          className="time24-dropdown"
          style={{ position: "fixed", top: dropPos.top, left: dropPos.left, zIndex: 9999 }}
        >
          <div className="time24-col" ref={hourRef}>
            {Array.from({ length: 24 }, (_, i) => (
              <button key={i} type="button"
                className={`time24-item${i === selH ? " time24-selected" : ""}`}
                onClick={() => pick(i, selM)}>
                {String(i).padStart(2, "0")}
              </button>
            ))}
          </div>
          <div className="time24-col" ref={minRef}>
            {Array.from({ length: 60 }, (_, i) => (
              <button key={i} type="button"
                className={`time24-item${i === selM ? " time24-selected" : ""}`}
                onClick={() => pick(selH, i, true)}>
                {String(i).padStart(2, "0")}
              </button>
            ))}
          </div>
        </div>,
        document.body,
      )
    : null;

  return (
    <>
      <button
        ref={btnRef}
        type="button"
        className={className ?? "time24-btn"}
        style={style}
        title={title}
        onClick={handleToggle}
      >
        <Clock size={13} className="time24-icon" />
        {value || "--:--"}
      </button>
      {dropdown}
    </>
  );
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

const cpfPublicHolidaysByYear: Record<string, Array<{ date: string; name: string }>> = {
  "2026": [
    { date: "2026-01-01", name: "วันขึ้นปีใหม่" },
    { date: "2026-03-03", name: "วันมาฆบูชา" },
    { date: "2026-04-06", name: "วันจักรี" },
    { date: "2026-04-13", name: "วันสงกรานต์" },
    { date: "2026-04-14", name: "วันสงกรานต์" },
    { date: "2026-04-15", name: "วันสงกรานต์" },
    { date: "2026-05-01", name: "วันแรงงานแห่งชาติ" },
    { date: "2026-05-04", name: "วันฉัตรมงคล" },
    { date: "2026-06-01", name: "วันหยุดชดเชยวันวิสาขบูชา" },
    { date: "2026-06-03", name: "วันเฉลิมพระชนมพรรษา สมเด็จพระนางเจ้าฯ พระบรมราชินี" },
    { date: "2026-07-28", name: "วันเฉลิมพระชนมพรรษา พระบาทสมเด็จพระเจ้าอยู่หัว" },
    { date: "2026-07-29", name: "วันอาสาฬหบูชา" },
    { date: "2026-07-30", name: "วันเข้าพรรษา" },
    { date: "2026-08-12", name: "วันแม่แห่งชาติ" },
    { date: "2026-10-13", name: "วันคล้ายวันสวรรคต ร.9" },
    { date: "2026-10-23", name: "วันปิยมหาราช" },
    { date: "2026-12-05", name: "วันพ่อแห่งชาติ" },
    { date: "2026-12-31", name: "วันสิ้นปี" },
  ],
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
  warnedIds?: Set<string>,
) {
  if (key === "minutesLate") return row.minutesLate;
  if (key === "monthlyLate") return monthlyLateCounts[row.empId] ?? 0;
  if (key === "warnCount") return warnCountMap[row.empId] ?? 0;
  if (key === "warned") return warnedIds?.has(row.empId) ? 1 : 0;
  return (row as Record<string, unknown>)[key] as string ?? "";
}

function sortAttendanceRows(
  rows: AttendanceRecord[],
  sort: SortState,
  monthlyLateCounts: Record<string, number> = {},
  warnCountMap: Record<string, number> = {},
  warnedIds?: Set<string>,
) {
  if (!sort) return rows;

  return [...rows].sort((a, b) => {
    const aValue = getAttendanceSortValue(a, sort.key, monthlyLateCounts, warnCountMap, warnedIds);
    const bValue = getAttendanceSortValue(b, sort.key, monthlyLateCounts, warnCountMap, warnedIds);

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
  defaultDirection = "asc",
}: {
  children: ReactNode;
  columnKey: string;
  setSort?: (sort: SortState) => void;
  sort?: SortState;
  defaultDirection?: SortDirection;
}) {
  const active = sort?.key === columnKey;
  return (
    <button
      className={`sort-button ${active ? "active" : ""}`}
      disabled={!setSort}
      onClick={() => {
        if (!setSort) return;
        if (!sort || sort.key !== columnKey) {
          setSort({ key: columnKey, direction: defaultDirection });
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

// Some exports use the Thai Buddhist Era (e.g. 2569) instead of Gregorian (2026).
function normalizeYear(year: number): number {
  return year > 2400 ? year - 543 : year;
}

function parseTimestamp(value: unknown) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return excelSerialDateToDate(value);
  }

  const text = String(value ?? "").trim();
  if (!text) return null;

  // ISO-style YYYY-MM-DD[ T]HH:MM[:SS] — unambiguous, always year-month-day.
  // Not anchored at the end so trailing content (milliseconds, "Z", a UTC offset, etc. — anything
  // a device/export might append) is tolerated instead of falling through to the ambiguous
  // native Date() parser below.
  const isoStyle = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (isoStyle) {
    const [, year, month, day, hour = "0", minute = "0", second = "0"] = isoStyle;
    const m = Number(month);
    const d = Number(day);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return new Date(normalizeYear(Number(year)), m - 1, d, Number(hour), Number(minute), Number(second));
    }
  }

  // DD-MM-YYYY / DD/MM/YYYY / DD.MM.YYYY [HH:MM[:SS]] — the format the time-attendance device
  // exports. Always day-month-year (never month-day): treating it as month-day is what silently
  // shifted every day 1-12 of every month to the wrong month (see the June/July timestamp bug).
  const dmyStyle = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (dmyStyle) {
    const [, day, month, year, hour = "0", minute = "0", second = "0"] = dmyStyle;
    const m = Number(month);
    const d = Number(day);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return new Date(normalizeYear(Number(year)), m - 1, d, Number(hour), Number(minute), Number(second));
    }
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
  const h = Number(match[1]);
  const m = Number(match[2]);
  if (h > 23 || m > 59) return "";
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

// ยึดหน่วยงานย่อยที่ตั้งไว้ในระบบอยู่แล้วก่อน (ล่าสุด) — ใช้ SKILL CF เป็นแค่ค่าเริ่มต้นสำหรับคนที่ยังไม่มีข้อมูล
function resolveCombinedJobSite(section: string | undefined, skillCF: string | undefined): string {
  return section || skillCF || "";
}

function excelSerialDateToDate(value: number) {
  // Use local midnight as base so .getHours()/.getDate() return correct local values.
  // Excel serial encodes wall-clock time, not UTC.
  const days = Math.floor(value);
  const ms = Math.round((value - days) * 86400000);
  const base = new Date(1899, 11, 30); // local midnight Dec 30 1899
  return new Date(base.getTime() + days * 86400000 + ms);
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
  let diff = Math.round((scanIn.getTime() - shift.getTime()) / 60000);
  // Night shifts: clock-in after midnight produces a large negative diff against
  // a same-date reference (e.g. 00:30 − 22:00 = −1290 min). Add one day to correct.
  if (diff < -720) diff += 1440;
  return diff;
}

// Returns the clock-out timestamp: the latest scan that is NOT the chosen clock-in.
//
// When scanIn is undefined (no clock-in found — early-morning scans belong to the
// previous night's shift), return undefined too. Those scans will be shown as
// scanOut on the PREVIOUS day's record (via nextDayScanRows augmentation there).
//
// Returns undefined when there is only one scan (clock-in only, no clock-out).
function findScanOut(scans: Date[], scanIn: Date | undefined): Date | undefined {
  if (!scans.length || !scanIn) return undefined;
  const after = scans.filter((t) => t.getTime() > scanIn.getTime());
  if (!after.length) return undefined;
  return after.sort((a, b) => b.getTime() - a.getTime())[0];
}

// Returns true when the time-of-day of `t` falls within the clock-in window:
// [shiftStart - 2 h, shiftStart + afterMin], wrapping around midnight for night shifts.
// Night shifts use +2 h (afterMin=120) to avoid misidentifying early-morning clock-outs
// (e.g. 02:05 AM exit) as clock-ins for a 22:00 or 23:00 shift.
function isInClockInWindow(t: Date, shiftHour: number, shiftMin: number, afterMin = 240): boolean {
  const tMin = t.getHours() * 60 + t.getMinutes();
  const shiftTotal = shiftHour * 60 + shiftMin;
  const windowStart = (shiftTotal - 120 + 1440) % 1440;
  const windowEnd = (shiftTotal + afterMin) % 1440;
  return windowStart <= windowEnd
    ? tMin >= windowStart && tMin <= windowEnd
    : tMin >= windowStart || tMin <= windowEnd; // window crosses midnight
}

// Pick the scan timestamp that is most likely the clock-in for `shiftStart`.
//
// Night shifts (shiftStart ≥ 20:00): search across all dates using the clock-in
// window only — the window (e.g. 20:00–02:00 for 22:00) naturally excludes the
// next-morning clock-out.
//
// Day/afternoon shifts: walk dates in descending order (most recent ≤ isoTargetDate
// first) and stop at the first date that has a timestamp inside the clock-in window.
// This handles two edge cases:
//   • overtime pushing the clock-out to the next calendar day (e.g. 01:00 on D+1)
//     — D+1 has no window match so we fall through to D and pick 13:57 correctly.
//   • multi-day scan files where previous-day clock-ins for the same shift would
//     otherwise beat today's by being slightly closer to shiftStart.
function findScanIn(scans: Date[], shiftStart: string, isoTargetDate: string): Date | undefined {
  if (!scans.length) return undefined;
  const [sh, sm] = shiftStart.split(":").map(Number);
  if (isNaN(sh) || isNaN(sm)) return scans.sort((a, b) => a.getTime() - b.getTime())[0];

  const shiftTotal = sh * 60 + sm;
  const toIso = (t: Date) =>
    `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;

  const pickClosest = (pool: Date[]) =>
    pool
      .map((t) => {
        const tMin = t.getHours() * 60 + t.getMinutes();
        let diff = Math.abs(tMin - shiftTotal);
        if (diff > 720) diff = 1440 - diff;
        return { t, diff };
      })
      .sort((a, b) => a.diff - b.diff)[0].t;

  // Shared helper: accept a previous-day candidate only when it is from exactly
  // the preceding calendar day. An overnight shift spans at most 1 day boundary,
  // so a scan from 2+ days ago is always stale data from a multi-day scan file.
  const isPrevDayOnly = (candidateIso: string): boolean => {
    const [cy, cm, cd] = candidateIso.split("-").map(Number);
    const [ty, tm, td] = isoTargetDate.split("-").map(Number);
    const diffMs = new Date(ty, tm - 1, td).getTime() - new Date(cy, cm - 1, cd).getTime();
    return Math.round(diffMs / 86400000) === 1;
  };

  if (sh >= 20) {
    // Night shift — use tight +2 h window to exclude early-morning clock-outs
    // (e.g. 02:05 AM exit for a 22:00 shift would fall inside a +4 h window).
    // If nothing lands in the window the available scans are early-morning
    // clock-outs from the previous night; return undefined so the caller marks
    // this day as Absent instead of misidentifying the exit scan as a clock-in.
    const inWindow = scans.filter((t) => isInClockInWindow(t, sh, sm, 120));
    if (!inWindow.length) return undefined;
    const candidate = pickClosest(inWindow);
    // Reject stale night-shift clock-ins from multi-day scan files (e.g. a 22:00
    // scan from 2 days ago would still pass the window check).
    if (isoTargetDate && toIso(candidate) < isoTargetDate && !isPrevDayOnly(toIso(candidate))) {
      return undefined;
    }
    return candidate;
  }

  // Day / afternoon shift — group by date and walk backwards.
  const byDate = new Map<string, Date[]>();
  for (const t of scans) {
    const d = toIso(t);
    byDate.set(d, [...(byDate.get(d) ?? []), t]);
  }
  const sortedDates = [...byDate.keys()]
    .filter((d) => !isoTargetDate || d <= isoTargetDate)
    .sort((a, b) => b.localeCompare(a)); // descending — most recent first

  for (const date of sortedDates) {
    const inWindow = (byDate.get(date) ?? []).filter((t) => isInClockInWindow(t, sh, sm));
    if (inWindow.length) {
      const candidate = pickClosest(inWindow);
      // If clock-in candidate is from a previous day, it is only valid when:
      // 1. It is from exactly the preceding calendar day (overnight shift, not 2+ days old).
      // 2. ALL of today's scans are before noon — the employee clocked out this morning,
      //    confirming a night-shift pattern (e.g. scanIn yesterday 17:00, scanOut today 01:30).
      if (isoTargetDate && toIso(candidate) < isoTargetDate) {
        if (!isPrevDayOnly(toIso(candidate))) return undefined;
        const todayScans = byDate.get(isoTargetDate) ?? [];
        const isNightShiftClockout = todayScans.length > 0 && todayScans.every((t) => t.getHours() < 12);
        if (!isNightShiftClockout) return undefined;
      }
      return candidate;
    }
  }

  // Fallback: earliest timestamp on or before isoTargetDate.
  const fallback = isoTargetDate ? scans.filter((t) => toIso(t) <= isoTargetDate) : scans;
  const candidate = (fallback.length ? fallback : scans).sort((a, b) => a.getTime() - b.getTime())[0];
  if (candidate && isoTargetDate && toIso(candidate) < isoTargetDate) {
    if (!isPrevDayOnly(toIso(candidate))) return undefined;
    const todayScans = byDate.get(isoTargetDate) ?? [];
    const isNightShiftClockout = todayScans.length > 0 && todayScans.every((t) => t.getHours() < 12);
    if (!isNightShiftClockout) return undefined;
  }
  // For PM shifts (shiftStart ≥ 12:00), an AM fallback scan is an exit from the
  // previous night's shift — treat it as absent rather than a very-late clock-in.
  if (candidate && sh >= 12 && candidate.getHours() < 12) {
    return undefined;
  }
  // A fallback scan more than 6 h outside the normal clock-in window is more likely
  // an OT exit than a late arrival — don't misidentify it as clock-in.
  if (candidate && !isInClockInWindow(candidate, sh, sm, 360)) {
    return undefined;
  }
  return candidate;
}

const STATUS_TH: Record<string, string> = {
  Present: "ตรงเวลา",
  Late: "มาสาย",
  Absent: "ขาดงาน",
  DayOff: "วันหยุด",
  Pending: "รอเข้างาน",
  NoScanIn: "ขาดสแกนเข้า",
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
      </div>
      <div className="kpi-stacked-bar">
        <div className="kpi-bar-fill present" style={{ width: `${presentPct}%` }} />
        <div className="kpi-bar-fill late" style={{ width: `${latePct}%` }} />
        <div className="kpi-bar-fill absent" style={{ width: `${absentPct}%` }} />
        <div className="kpi-bar-fill dayoff" style={{ width: `${dayoffPct}%` }} />
      </div>
      <div className="kpi-donut-legend">
        <LegendRow color="green" label="ตรงเวลา" value={String(present)} percent={`${presentPct.toFixed(1)}%`} />
        <LegendRow color="amber" label="มาสาย" value={String(late)} percent={`${latePct.toFixed(1)}%`} />
        <LegendRow color="red" label="ขาด/ลา" value={String(absent)} percent={`${absentPct.toFixed(1)}%`} />
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
    .filter((r) => r.status === "Late" || r.status === "Absent" || r.status === "NoScanIn" || r.status === "Pending")
    .sort((a, b) => {
      const order: Record<string, number> = { Absent: 0, Pending: 1, NoScanIn: 2, Late: 3 };
      const ao = order[a.status] ?? 9, bo = order[b.status] ?? 9;
      if (ao !== bo) return ao - bo;
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
      "Scan In": r.scanInDate ? `${r.scanInDate} ${r.scanIn}` : r.scanIn,
      "Scan Out": r.scanOutDate ? `${r.scanOutDate} ${r.scanOut}` : r.scanOut,
      "สถานะ": STATUS_TH[r.status] ?? r.status,
      "สาย (นาที)": r.minutesLate,
      "สายสะสมเดือนนี้ (ครั้ง)": monthlyLateCounts[r.empId] ?? 0,
    }));

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Late-Absent");
  const today = new Date();
  const dateStr = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
  const filename = `late-absent_${deptLabel}_${dateStr}.xlsx`;
  XLSX.writeFile(wb, filename);
}

function DashboardPanels({
  activeMasterMap,
  assignedPeople,
  dashboardDeptFilter,
  dashboardSectionFilter,
  detailStatusFilter,
  setDetailStatusFilter,
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
  dashboardSectionFilter: string;
  detailStatusFilter: string;
  setDetailStatusFilter: (v: string) => void;
  isoTargetDate: string;
  reportData: ReportData | null;
  toggleWarning: (empId: string) => Promise<void>;
  totalActivePeople: number;
  warnCountMap: Record<string, number>;
  warnedIds: Set<string>;
  warnPending: Set<string>;
}) {
  const [detailSort, setDetailSort_] = useState<SortState>(null);
  const setDetailSort = setDetailSort_ as (sort: SortState) => void;
  const [lateSort, setLateSort_] = useState<SortState>(null);
  const setLateSort = setLateSort_ as (sort: SortState) => void;
  const [leaveMap, setLeaveMap] = useState<Map<string, string>>(new Map());
  const [leaveError, setLeaveError] = useState("");
  const [warnPanelCollapsed, setWarnPanelCollapsed] = useState(true);

  useEffect(() => {
    if (!isoTargetDate) return;
    let cancelled = false;
    setLeaveError("");
    supabase.from("leave_records").select("emp_id, leave_type").eq("leave_date", isoTargetDate)
      .then(({ data: rows, error }) => {
        if (cancelled) return;
        if (error) { setLeaveError(`โหลดข้อมูลการลาไม่สำเร็จ: ${error.message}`); return; }
        const map = new Map((rows ?? []).map((r: { emp_id: string; leave_type: string }) => [r.emp_id, r.leave_type]));
        const absentIds = (reportData?.records ?? [])
          .filter((r: { status: string }) => r.status === "Absent")
          .map((r: { empId: string }) => r.empId);
        absentIds.filter((id: string) => !map.has(id)).forEach((id: string) => map.set(id, "ขาดงาน"));
        setLeaveMap(map);
      });
    return () => { cancelled = true; };
  }, [isoTargetDate, reportData]); // eslint-disable-line react-hooks/exhaustive-deps

  const saveLeave = async (empId: string, leaveType: string) => {
    if (!isoTargetDate || !leaveType) return;
    const prev = leaveMap.get(empId);
    setLeaveMap(m => new Map(m).set(empId, leaveType));
    const { error } = await supabase.from("leave_records").upsert(
      { emp_id: empId, leave_date: isoTargetDate, leave_type: leaveType },
      { onConflict: "emp_id,leave_date" }
    );
    if (error) {
      setLeaveMap(m => { const n = new Map(m); prev === undefined ? n.delete(empId) : n.set(empId, prev); return n; });
    }
  };

  const [confirmation, setConfirmation] = useState<{ confirmed_by: string; confirmed_at: string } | null | undefined>(undefined);
  const [confirmName, setConfirmName] = useState("");
  const [isConfirming, setIsConfirming] = useState(false);

  useEffect(() => {
    if (!isoTargetDate) { setConfirmation(undefined); setConfirmName(""); return; }
    let cancelled = false;
    setConfirmation(undefined);
    setConfirmName("");
    let deptKey = dashboardDeptFilter === "all" ? "ทุกหน่วยงาน" : dashboardDeptFilter;
    if (dashboardSectionFilter !== "all") deptKey = `${deptKey}/${dashboardSectionFilter}`;
    supabase.from("daily_confirmations")
      .select("confirmed_by, confirmed_at")
      .eq("confirm_date", isoTargetDate)
      .eq("dept", deptKey)
      .maybeSingle()
      .then(({ data, error }) => {
        if (cancelled) return;
        if (error) { console.error("daily_confirmations load error:", error.message); setConfirmation(null); return; }
        setConfirmation(data ?? null);
      });
    return () => { cancelled = true; };
  }, [isoTargetDate, dashboardDeptFilter, dashboardSectionFilter]);

  const handleConfirm = async () => {
    if (isConfirming) return;
    if (!confirmName.trim()) {
      const input = document.querySelector<HTMLInputElement>(".confirm-name-input");
      input?.focus();
      input?.classList.add("confirm-input-error");
      setTimeout(() => input?.classList.remove("confirm-input-error"), 1200);
      return;
    }
    if (!isoTargetDate) return;
    setIsConfirming(true);
    try {
      let deptKey = dashboardDeptFilter === "all" ? "ทุกหน่วยงาน" : dashboardDeptFilter;
      if (dashboardSectionFilter !== "all") deptKey = `${deptKey}/${dashboardSectionFilter}`;
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
    } finally {
      setIsConfirming(false);
    }
  };

  const total = reportData?.totalEmployees ?? 0;
  const present = reportData?.present ?? 0;
  const late = reportData?.late ?? 0;
  const absent = reportData?.absent ?? 0;
  const presentPct = total ? (present / total) * 100 : 0;
  const latePct = total ? (late / total) * 100 : 0;
  const absentPct = total ? (absent / total) * 100 : 0;
  const topDeptRows = reportData?.deptRows ?? [];
  const maxDeptTotal = Math.max(...topDeptRows.map((row) => row.total), 1);
  const monthlyLateCounts = reportData?.monthlyLateCounts ?? {};
  const dashboardLateRows = lateSort
    ? lateSort.key === "scanIn"
      ? [...(reportData?.lateRows ?? [])].sort((a, b) => {
          const aKey = `${a.scanInDate || isoTargetDate}T${a.scanIn}`;
          const bKey = `${b.scanInDate || isoTargetDate}T${b.scanIn}`;
          return lateSort.direction === "asc" ? aKey.localeCompare(bKey) : bKey.localeCompare(aKey);
        })
      : sortAttendanceRows(reportData?.lateRows ?? [], lateSort, monthlyLateCounts, warnCountMap, warnedIds)
    : [...(reportData?.lateRows ?? [])].sort((a, b) => {
        const aDate = a.scanInDate || isoTargetDate;
        const bDate = b.scanInDate || isoTargetDate;
        if (aDate !== bDate) return bDate.localeCompare(aDate);
        return b.scanIn.localeCompare(a.scanIn);
      });
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
      {leaveError && (
        <div className="error-banner" style={{ marginBottom: 8 }}>{leaveError}</div>
      )}
      {isoTargetDate && (
        <section className="confirm-attendance-card">
          <div className="confirm-card-top">
            <div className="confirm-card-title-row">
              <ClipboardCheck size={15} className="confirm-title-icon" />
              <span className="confirm-title-text">ยืนยันตรวจสอบการเข้างาน</span>
              <span className="confirm-title-dept">{dashboardDeptFilter === "all" ? "ทุกหน่วยงาน" : dashboardDeptFilter}</span>
            </div>
            <div className="confirm-summary-row">
              <span className="csb present" style={{ cursor: "pointer" }} onClick={() => { setDetailStatusFilter("Present"); setTimeout(() => document.getElementById("employee-detail-section")?.scrollIntoView({ behavior: "smooth", block: "start" }), 50); }}>ตรงเวลา {present}</span>
              <span className="csb late" style={{ cursor: "pointer" }} onClick={() => { setDetailStatusFilter("Late"); setTimeout(() => document.getElementById("employee-detail-section")?.scrollIntoView({ behavior: "smooth", block: "start" }), 50); }}>มาสาย {late}</span>
              <div className="csb-absent-group" style={{ cursor: "pointer" }} onClick={() => { setDetailStatusFilter("Absent"); setTimeout(() => document.getElementById("employee-detail-section")?.scrollIntoView({ behavior: "smooth", block: "start" }), 50); }}>
                <span className="csb absent">ขาด/ลา {absent}</span>
                {absent > 0 && (() => {
                  const deptAbsentIds = new Set(allRecords.filter(r => r.status === "Absent").map(r => r.empId));
                  const leaveColor: Record<string, string> = { "ลาป่วย": "blue", "ลากิจ": "amber", "ลาพักร้อน": "green", "ขาดงาน": "red", "ลาตรวจครรภ์": "pink", "ลาคลอด": "rose", "ลาอุบัติเหตุจากการปฏิบัติงาน": "orange", "ลาบวช/ลาพิธีสำคัญทางศาสนา": "purple", "ลาทหาร": "indigo", "ลาพิเศษไม่จ่าย": "gray" };
                  const subBadges = (["ลาป่วย", "ลากิจ", "ลาพักร้อน", "ลาตรวจครรภ์", "ลาคลอด", "ลาอุบัติเหตุจากการปฏิบัติงาน", "ลาบวช/ลาพิธีสำคัญทางศาสนา", "ลาทหาร", "ลาพิเศษไม่จ่าย", "ขาดงาน"] as const).map(type => {
                    const cnt = [...leaveMap.entries()].filter(([eid, lt]) => deptAbsentIds.has(eid) && lt === type).length;
                    return cnt ? <span key={type} className={`csb-leave ${leaveColor[type] ?? "blue"}`}>{type} {cnt}</span> : null;
                  }).filter(Boolean);
                  return subBadges.length ? <><span className="csb-absent-divider" />{subBadges}</> : null;
                })()}
              </div>
            </div>
          </div>
          <div className="confirm-card-action">
            {confirmation === undefined ? (
              <span className="confirm-loading">กำลังโหลด...</span>
            ) : confirmation ? (
              <div className="confirm-done-row">
                <CheckCircle2 size={15} className="confirm-done-icon" />
                <span>รับทราบแล้วโดย <strong>{confirmation.confirmed_by}</strong> · {new Date(confirmation.confirmed_at).toLocaleString("th-TH", { day: "numeric", month: "numeric", year: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false })}</span>
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
        <section id="late-people-section" className="panel dashboard-late-card">
          <div className="panel-title-row">
            <h3>คนที่มาสาย</h3>
            <span className="table-count">{dashboardLateRows.length} คน</span>
          </div>
          <div className="late-preview-table">
            <table className="table compact-table">
              <thead>
                <tr>
                  <th><SortButton columnKey="name" setSort={setLateSort} sort={lateSort} defaultDirection="desc">ชื่อ</SortButton></th>
                  <th><SortButton columnKey="dept" setSort={setLateSort} sort={lateSort} defaultDirection="desc">หน่วยงาน</SortButton></th>
                  <th><SortButton columnKey="shiftStart" setSort={setLateSort} sort={lateSort} defaultDirection="desc">เริ่มกะ</SortButton></th>
                  <th><SortButton columnKey="scanIn" setSort={setLateSort} sort={lateSort} defaultDirection="desc">เข้างาน</SortButton></th>
                  <th><SortButton columnKey="minutesLate" setSort={setLateSort} sort={lateSort} defaultDirection="desc">สาย</SortButton></th>
                  <th><SortButton columnKey="warnCount" setSort={setLateSort} sort={lateSort} defaultDirection="desc">เตือนสะสม</SortButton></th>
                  <th><SortButton columnKey="warned" setSort={setLateSort} sort={lateSort} defaultDirection="desc">เตือน</SortButton></th>
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
                    <td><span className="shift-start-badge">{row.shiftStart}</span></td>
                    <td className="scan-cell">{scanDateBadge(row.scanInDate || isoTargetDate)}{row.scanIn}</td>
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
                  <tr><td colSpan={7}>ยังไม่มีข้อมูลคนมาสาย</td></tr>
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
      <section id="employee-detail-section" className="panel detail-attendance-panel">
        <div className="panel-title-row">
          <h3>
            สถานะพนักงานรายคน
            {dashboardDeptFilter !== "all" ? ` · ${dashboardDeptFilter}` : ""}
          </h3>
        </div>
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
              <option value="DayOff">วันหยุด</option>
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
                    <td className="scan-cell">{scanDateBadge(row.scanInDate)}{row.scanIn}</td>
                    <td>
                      {row.status === "Absent" ? (() => {
                        const lt = leaveMap.get(row.empId) ?? "ขาดงาน";
                        const lc = ({ "ลาป่วย": "leave-sick", "ลากิจ": "leave-personal", "ลาพักร้อน": "leave-vacation", "ลาตรวจครรภ์": "leave-prenatal", "ลาคลอด": "leave-maternity", "ลาอุบัติเหตุจากการปฏิบัติงาน": "leave-accident", "ลาบวช/ลาพิธีสำคัญทางศาสนา": "leave-ordain", "ลาทหาร": "leave-military", "ลาพิเศษไม่จ่าย": "leave-unpaid" } as Record<string, string>)[lt] ?? "leave-absent";
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
                            <option value="ลาตรวจครรภ์">ลาตรวจครรภ์</option>
                            <option value="ลาคลอด">ลาคลอด</option>
                            <option value="ลาอุบัติเหตุจากการปฏิบัติงาน">ลาอุบัติเหตุจากการปฏิบัติงาน</option>
                            <option value="ลาบวช/ลาพิธีสำคัญทางศาสนา">ลาบวช/ลาพิธีสำคัญทางศาสนา</option>
                            <option value="ลาทหาร">ลาทหาร</option>
                            <option value="ลาพิเศษไม่จ่าย">ลาพิเศษไม่จ่าย</option>
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
      </section>

    </>
  );
}

function buildCombinedExcelWorkbook(
  empRows: Record<string, unknown>[],
  dayoffMap: Map<string, { dayoff: string; shift: string; shiftStart: string; shiftEnd: string; section: string }>,
  skillFlatRows: SkillFlatRow[],
  manpowerRows: Record<string, unknown>[],
  skillNames: string[],
  skillCFMap: Map<string, string>,
): ReturnType<typeof XLSX.utils.book_new> {
  const sheet1Rows = empRows.map((row) => {
    const empId = cleanEmpId(row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"]);
    if (!empId) return null;
    const ds = dayoffMap.get(empId);
    const empSkills = skillFlatRows.filter((s) => s.empId === empId);
    const skillCols: Record<string, number> = {};
    for (const name of skillNames) skillCols[name] = empSkills.find((s) => s.skill === name)?.level ?? 0;
    return {
      "Employee ID": empId,
      "First Name (Local)": String(row["First Name (Local)"] ?? "").trim(),
      "Last Name (Local)": String(row["Last Name (Local)"] ?? "").trim(),
      "หน่วยงาน": String(row["หน่วยงาน"] ?? row["Name (Section)"] ?? "").trim(),
      "Title (Position)": String(row["Title (Position)"] ?? row["position"] ?? "").trim(),
      "หน่วยงานย่อย/Skill": resolveCombinedJobSite(ds?.section, skillCFMap.get(empId)),
      "กะ": ds?.shift ?? "",
      "เวลาเข้างาน": ds?.shiftStart ?? "",
      "เวลาออก": ds?.shiftEnd ?? "",
      "วันหยุดประจำสัปดาห์": ds?.dayoff ?? "",
      ...skillCols,
    };
  }).filter(Boolean);
  const ws1 = XLSX.utils.json_to_sheet(sheet1Rows as object[]);
  ws1["!cols"] = [
    { wch: 14 }, { wch: 18 }, { wch: 18 }, { wch: 22 }, { wch: 20 }, { wch: 16 },
    { wch: 8 }, { wch: 14 }, { wch: 14 }, { wch: 24 },
    ...skillNames.map(() => ({ wch: 14 })),
  ];
  const ws2 = XLSX.utils.json_to_sheet(
    manpowerRows.length ? manpowerRows : [{ "หน่วยงาน": "", "กะ": "", "เวลาเข้า": "" }],
  );
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "พนักงาน");
  XLSX.utils.book_append_sheet(wb, ws2, "Manpower Plan");
  return wb;
}

function parseCombinedSheet1(rows: Record<string, unknown>[]): CombinedEmployeeRow[] {
  const FIXED = new Set([
    "employee id", "first name (local)", "last name (local)",
    "หน่วยงาน", "title (position)", "หน่วยงานย่อย/skill", "หน้างาน", "กะ", "เวลาเข้างาน", "เวลาออก", "วันหยุดประจำสัปดาห์",
  ]);
  const norm = (s: string) => s.toLowerCase().replace(/\s+/g, " ").trim();
  if (rows.length > 0) {
    const hasEmpIdCol = Object.keys(rows[0]).some((k) => norm(k) === "employee id");
    if (!hasEmpIdCol) throw new Error("ไม่พบคอลัมน์ 'Employee ID' ในไฟล์ — กรุณาใช้ไฟล์ที่ Export จากระบบนี้");
  }
  return rows.map((row) => {
    const empId = cleanEmpId(row["Employee ID"] ?? "");
    if (!empId) return null;
    const skills: Record<string, number> = {};
    for (const [key, val] of Object.entries(row)) {
      if (!FIXED.has(norm(key)) && key.trim()) {
        const level = Number(val);
        if (!isNaN(level)) skills[key.trim()] = level;
      }
    }
    return {
      empId,
      firstName: String(row["First Name (Local)"] ?? "").trim(),
      lastName: String(row["Last Name (Local)"] ?? "").trim(),
      dept: String(row["หน่วยงาน"] ?? "").trim(),
      position: String(row["Title (Position)"] ?? "").trim(),
      jobSite: String(row["หน่วยงานย่อย/Skill"] ?? row["หน้างาน"] ?? "").trim(),
      shift: String(row["กะ"] ?? "").trim(),
      shiftStart: normalizeTimeText(row["เวลาเข้างาน"]),
      shiftEnd: normalizeTimeText(row["เวลาออก"]),
      dayoff: String(row["วันหยุดประจำสัปดาห์"] ?? "").trim(),
      skills,
    };
  }).filter(Boolean) as CombinedEmployeeRow[];
}

function computeEmployeeDiff(
  newRows: CombinedEmployeeRow[],
  currentRows: CombinedEmployeeRow[],
  newManpowerRows: Record<string, unknown>[],
): EmployeeDiff {
  const curMap = new Map(currentRows.map((r) => [r.empId, r]));
  const newMap = new Map(newRows.map((r) => [r.empId, r]));
  const added = newRows.filter((r) => !curMap.has(r.empId));
  const removed = currentRows.filter((r) => !newMap.has(r.empId));
  const changed: EmployeeDiff["changed"] = [];
  let unchangedCount = 0;
  for (const nr of newRows) {
    const cur = curMap.get(nr.empId);
    if (!cur) continue;
    const fields: Array<{ field: string; from: string; to: string }> = [];
    const chk = (field: string, from: string, to: string) => {
      if (from.trim() !== to.trim()) fields.push({ field, from, to });
    };
    chk("ชื่อ", `${cur.firstName} ${cur.lastName}`.trim(), `${nr.firstName} ${nr.lastName}`.trim());
    chk("หน่วยงาน", cur.dept, nr.dept);
    chk("ตำแหน่ง", cur.position, nr.position);
    chk("หน่วยงานย่อย/Skill", cur.jobSite, nr.jobSite);
    chk("กะ", cur.shift, nr.shift);
    chk("เวลาเข้างาน", cur.shiftStart, nr.shiftStart);
    chk("เวลาออก", cur.shiftEnd, nr.shiftEnd);
    chk("วันหยุด", cur.dayoff, nr.dayoff);
    const allSkills = new Set([...Object.keys(cur.skills), ...Object.keys(nr.skills)]);
    for (const s of allSkills) {
      if (!s.trim()) continue;
      const fl = cur.skills[s] ?? 0;
      const tl = nr.skills[s] ?? 0;
      if (fl !== tl) chk(`ทักษะ: ${s}`, String(fl), String(tl));
    }
    if (fields.length) changed.push({ empId: nr.empId, name: `${nr.firstName} ${nr.lastName}`.trim(), dept: nr.dept, fields });
    else unchangedCount++;
  }
  return { added, removed, changed, unchangedCount, newRows, newManpowerRows };
}

function DiffPreviewModal({
  diff,
  onConfirm,
  onCancel,
  isSaving,
}: {
  diff: EmployeeDiff;
  onConfirm: () => void;
  onCancel: () => void;
  isSaving: boolean;
}) {
  const [deletionConfirmed, setDeletionConfirmed] = useState(false);
  const canConfirm = diff.removed.length === 0 || deletionConfirmed;
  const defaultSection = diff.removed.length > 0 ? "removed" : diff.added.length > 0 ? "added" : diff.changed.length > 0 ? "changed" : null;
  const [expandSection, setExpandSection] = useState<"added" | "removed" | "changed" | null>(defaultSection);
  const toggle = (s: "added" | "removed" | "changed") => setExpandSection((prev) => (prev === s ? null : s));

  return (
    <div className="modal-overlay">
      <div className="modal-panel diff-modal">
        <div className="modal-header">
          <h3>ยืนยันการอัพเดทข้อมูลพนักงาน</h3>
          <button className="icon-button" onClick={onCancel} disabled={isSaving} type="button">
            <X size={18} />
          </button>
        </div>

        <div className="diff-summary-row">
          {diff.added.length > 0 && (
            <button className={`diff-badge added${expandSection === "added" ? " active" : ""}`} onClick={() => toggle("added")} type="button">
              + เพิ่ม {diff.added.length} คน
            </button>
          )}
          {diff.removed.length > 0 && (
            <button className={`diff-badge removed${expandSection === "removed" ? " active" : ""}`} onClick={() => toggle("removed")} type="button">
              − ลบ {diff.removed.length} คน
            </button>
          )}
          {diff.changed.length > 0 && (
            <button className={`diff-badge changed${expandSection === "changed" ? " active" : ""}`} onClick={() => toggle("changed")} type="button">
              ≈ แก้ไข {diff.changed.length} คน
            </button>
          )}
          <span className="diff-badge unchanged">= ไม่เปลี่ยน {diff.unchangedCount} คน</span>
        </div>

        <div className="diff-body">
          {expandSection === "added" && (
            <div className="diff-section added">
              <div className="diff-section-title">เพิ่มใหม่</div>
              {diff.added.map((r) => (
                <div key={r.empId} className="diff-row">
                  <span className="diff-empid">{r.empId}</span>
                  <span className="diff-name">{`${r.firstName} ${r.lastName}`.trim() || "—"}</span>
                  <span className="dept-chip">{r.dept || "—"}</span>
                </div>
              ))}
            </div>
          )}
          {expandSection === "removed" && (
            <div className="diff-section removed">
              <div className="diff-section-title">ลบออก</div>
              {diff.removed.map((r) => (
                <div key={r.empId} className="diff-row">
                  <span className="diff-empid">{r.empId}</span>
                  <span className="diff-name">{`${r.firstName} ${r.lastName}`.trim() || "—"}</span>
                  <span className="dept-chip">{r.dept || "—"}</span>
                </div>
              ))}
              <label className="diff-delete-confirm">
                <input type="checkbox" checked={deletionConfirmed} onChange={(e) => setDeletionConfirmed(e.target.checked)} />
                ฉันยืนยันว่าต้องการลบพนักงาน {diff.removed.length} คนนี้ออกจากระบบ
              </label>
            </div>
          )}
          {expandSection === "changed" && (
            <div className="diff-section changed">
              <div className="diff-section-title">เปลี่ยนแปลง</div>
              {diff.changed.slice(0, 30).map((r) => (
                <div key={r.empId} className="diff-changed-row">
                  <div className="diff-row">
                    <span className="diff-empid">{r.empId}</span>
                    <span className="diff-name">{r.name || "—"}</span>
                    <span className="dept-chip">{r.dept || "—"}</span>
                  </div>
                  {r.fields.map((f) => (
                    <div key={f.field} className="diff-field-row">
                      <span className="diff-field-label">{f.field}</span>
                      <span className="diff-from">{f.from || "(ว่าง)"}</span>
                      <span className="diff-arrow">→</span>
                      <span className="diff-to">{f.to || "(ว่าง)"}</span>
                    </div>
                  ))}
                </div>
              ))}
              {diff.changed.length > 30 && (
                <p style={{ color: "var(--muted)", fontSize: 13, margin: "8px 0 0" }}>
                  ...และอีก {diff.changed.length - 30} คน
                </p>
              )}
            </div>
          )}
          {diff.added.length === 0 && diff.removed.length === 0 && diff.changed.length === 0 && (
            <p style={{ textAlign: "center", color: "var(--muted)", fontSize: 14, padding: "24px 0" }}>
              ข้อมูลไม่มีการเปลี่ยนแปลง
            </p>
          )}
        </div>

        <div className="modal-footer">
          <button className="secondary-button" onClick={onCancel} disabled={isSaving} type="button">
            ยกเลิก
          </button>
          <button className="primary-button" disabled={!canConfirm || isSaving} onClick={onConfirm} type="button">
            <UploadCloud size={16} />
            {isSaving ? "กำลังบันทึก..." : `ยืนยันอัพเดท → ${diff.newRows.length} คน`}
          </button>
        </div>
      </div>
    </div>
  );
}

function MasterDataPage({
  activeMasterMap,
  isSavingMasters,
  masterFileHistory,
  masterSubTab,
  masterUploads,
  onDeleteMasterFile,
  onHolidaysChanged,
  onMastersSaved,
  saveDayoffShiftRows,
  saveManpowerRows,
  saveMasterFiles,
  setError,
  setMessage,
  setMasterUploads,
}: {
  activeMasterMap: Partial<Record<MasterFileKey, MasterFile>>;
  isSavingMasters: boolean;
  masterFileHistory: MasterFile[];
  masterSubTab: "files" | "holidays" | "public_holidays" | "manpower" | "dayoff_shift";
  masterUploads: MasterUploadState;
  onDeleteMasterFile: (file: MasterFile) => Promise<void>;
  onHolidaysChanged: (dates: Set<string>) => void;
  onMastersSaved: () => Promise<void>;
  saveDayoffShiftRows: (rows: DayoffShiftEditorRow[]) => Promise<void>;
  saveManpowerRows: (rows: ManpowerEditorRow[]) => Promise<void>;
  saveMasterFiles: () => Promise<void>;
  setError: (msg: string) => void;
  setMessage: (msg: string) => void;
  setMasterUploads: Dispatch<SetStateAction<MasterUploadState>>;
}) {
  const [combinedFile, setCombinedFile] = useState<File | null>(null);
  const [isSavingCombined, setIsSavingCombined] = useState(false);
  const [diffResult, setDiffResult] = useState<EmployeeDiff | null>(null);
  const [showDiffModal, setShowDiffModal] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  function downloadCombinedTemplate() {
    const headers = [
      "Employee ID", "First Name (Local)", "Last Name (Local)",
      "หน่วยงาน", "Title (Position)", "หน่วยงานย่อย/Skill", "กะ", "เวลาเข้างาน", "วันหยุดประจำสัปดาห์",
    ];
    const examples = [
      ["EMP001", "สมชาย", "ใจดี", "งานเครื่องใน", "พนักงานผลิต", "ตะกร้า", "กะ 1", "07:00", "อาทิตย์"],
      ["EMP002", "สมหญิง", "รักดี", "งานแยกชิ้นส่วน", "พนักงานผลิต", "ดันหมู", "กะ 1", "07:00", "เสาร์-อาทิตย์"],
      ["EMP003", "มานพ", "สุขใจ", "งานควบคุมคุณภาพ", "หัวหน้างาน", "QC", "กะ 2", "13:00", "อาทิตย์"],
    ];
    const ws = XLSX.utils.aoa_to_sheet([headers, ...examples]);
    ws["!cols"] = [14, 18, 18, 22, 20, 16, 8, 14, 24].map((w) => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "พนักงาน");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["หน่วยงาน", "กะ", "เวลาเข้า", "เวลาออก"], ["งานเครื่องใน", "กะ 1", "07:00", "16:00"]]), "Manpower Plan");
    XLSX.writeFile(wb, "template-master-พนักงาน.xlsx");
  }

  async function exportCombinedMaster() {
    setIsExporting(true);
    setError("");
    try {
      const skillCFPromise = (async () => {
        try {
          const res = await fetch("/skillcf_default.xlsx");
          if (!res.ok) return new Map<string, string>();
          const buffer = await res.arrayBuffer();
          const wb = XLSX.read(buffer, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const rawRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
          const hdrIdx = rawRows.findIndex((r) => (r as string[]).includes("User ID (Job Information)"));
          if (hdrIdx < 0) return new Map<string, string>();
          const hdr = rawRows[hdrIdx] as string[];
          const map = new Map<string, string>();
          for (const r of rawRows.slice(hdrIdx + 1)) {
            const row: Record<string, unknown> = {};
            (r as unknown[]).forEach((val, i) => { if (hdr[i]) row[hdr[i]] = val; });
            const empId = cleanEmpId(row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"]);
            const skillCF = String(row["SKILL CF"] ?? "").trim();
            if (empId && skillCF) map.set(empId, skillCF);
          }
          return map;
        } catch { return new Map<string, string>(); }
      })();
      const [empRows, dayoffRows, skillRawRows, manpowerRows, skillCFMap] = await Promise.all([
        activeMasterMap.employee_master ? downloadSheetRows(activeMasterMap.employee_master.file_path) : Promise.resolve([]),
        activeMasterMap.dayoff_shift ? downloadSheetRows(activeMasterMap.dayoff_shift.file_path) : Promise.resolve([]),
        activeMasterMap.skill_matrix ? downloadSheetRows(activeMasterMap.skill_matrix.file_path) : Promise.resolve([]),
        activeMasterMap.manpower_plan ? downloadSheetRows(activeMasterMap.manpower_plan.file_path) : Promise.resolve([]),
        skillCFPromise,
      ]);
      const skillFlatRows: SkillFlatRow[] = skillRawRows
        .map((row, i) => {
          const empId = cleanEmpId(row["Employee ID"] ?? row["Emp ID"] ?? "");
          const skill = String(row["Skill"] ?? row["skill"] ?? "").trim();
          const level = Number(row["Level"] ?? row["level"]) || 0;
          return { id: `${i}-${empId}-${skill}`, empId, name: "", dept: "", jobSite: "", shift: "", shiftStart: "", dayoff: "", skill, level, origLevel: level };
        })
        .filter((r) => r.empId && r.skill);
      const skillNames = Array.from(new Set(skillFlatRows.map((r) => r.skill))).sort();
      const dayoffMap = buildDayoffShiftMap(dayoffRows);
      const skippedCount = empRows.filter(
        (row) => !cleanEmpId(row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"]),
      ).length;
      const wb = buildCombinedExcelWorkbook(empRows, dayoffMap, skillFlatRows, manpowerRows, skillNames, skillCFMap);
      const now = new Date().toLocaleDateString("th-TH").replace(/\//g, "-");
      XLSX.writeFile(wb, `master-พนักงาน-${now}.xlsx`);
      if (skippedCount > 0) setMessage(`Export สำเร็จ — พนักงาน ${skippedCount} คนถูกข้ามเพราะไม่มี Employee ID`);
    } catch (e) {
      setError(e instanceof Error ? e.message : "Export ไม่สำเร็จ");
    } finally {
      setIsExporting(false);
    }
  }

  async function previewCombinedUpload() {
    if (!combinedFile) return;
    setIsSavingCombined(true);
    setError("");
    try {
      const ext = combinedFile.name.split(".").pop()?.toLowerCase() ?? "";
      if (!["xlsx", "xls"].includes(ext)) {
        setError("รองรับเฉพาะไฟล์ .xlsx หรือ .xls เท่านั้น");
        setIsSavingCombined(false);
        return;
      }
      if (combinedFile.size > 50 * 1024 * 1024) {
        setError("ไฟล์มีขนาดใหญ่เกินไป — สูงสุด 50 MB");
        setIsSavingCombined(false);
        return;
      }
      const buffer = await combinedFile.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array", cellDates: true });
      const sheet1Name = wb.SheetNames.find((n) => n.includes("พนักงาน") || n.toLowerCase().includes("employee")) ?? wb.SheetNames[0];
      const sheet1Rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(wb.Sheets[sheet1Name], { defval: "" });
      const newCombinedRows = parseCombinedSheet1(sheet1Rows);
      if (newCombinedRows.length === 0) {
        setError("ไม่พบข้อมูลพนักงานในไฟล์ — ตรวจสอบว่าใช้ไฟล์ที่ Export จากระบบนี้ และ Sheet แรกมีคอลัมน์ Employee ID");
        setIsSavingCombined(false);
        return;
      }
      const empIdCounts = new Map<string, number>();
      for (const r of newCombinedRows) empIdCounts.set(r.empId, (empIdCounts.get(r.empId) ?? 0) + 1);
      const dupes = [...empIdCounts.entries()].filter(([, n]) => n > 1).map(([id]) => id);
      if (dupes.length > 0) {
        setError(`พบ Employee ID ซ้ำในไฟล์: ${dupes.slice(0, 5).join(", ")}${dupes.length > 5 ? ` และอีก ${dupes.length - 5} รายการ` : ""} — กรุณาตรวจสอบและแก้ไขก่อนอัพโหลด`);
        setIsSavingCombined(false);
        return;
      }

      const sheet2Name = wb.SheetNames.find((n) => n.toLowerCase().includes("manpower") || n.toLowerCase().includes("plan")) ?? wb.SheetNames[1];
      const newManpowerRows = sheet2Name && wb.Sheets[sheet2Name]
        ? XLSX.utils.sheet_to_json<Record<string, unknown>>(wb.Sheets[sheet2Name], { defval: "" })
        : [];

      const loadSkillCFMap = async (): Promise<Map<string, string>> => {
        try {
          const res = await fetch("/skillcf_default.xlsx");
          if (!res.ok) return new Map<string, string>();
          const buffer = await res.arrayBuffer();
          const wb = XLSX.read(buffer, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const rawRows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "" });
          const hdrIdx = rawRows.findIndex((r) => (r as string[]).includes("User ID (Job Information)"));
          if (hdrIdx < 0) return new Map<string, string>();
          const hdr = rawRows[hdrIdx] as string[];
          const map = new Map<string, string>();
          for (const r of rawRows.slice(hdrIdx + 1)) {
            const rowObj: Record<string, unknown> = {};
            (r as unknown[]).forEach((val, i) => { if (hdr[i]) rowObj[hdr[i]] = val; });
            const empId = cleanEmpId(rowObj["User ID (Job Information)"] ?? rowObj["Employee ID"] ?? rowObj["Emp ID"]);
            const skillCF = String(rowObj["SKILL CF"] ?? "").trim();
            if (empId && skillCF) map.set(empId, skillCF);
          }
          return map;
        } catch { return new Map<string, string>(); }
      };
      const [curEmpRows, curDayoffRows, curSkillRows, curSkillCFMap] = await Promise.all([
        activeMasterMap.employee_master ? downloadSheetRows(activeMasterMap.employee_master.file_path) : Promise.resolve([]),
        activeMasterMap.dayoff_shift ? downloadSheetRows(activeMasterMap.dayoff_shift.file_path) : Promise.resolve([]),
        activeMasterMap.skill_matrix ? downloadSheetRows(activeMasterMap.skill_matrix.file_path) : Promise.resolve([]),
        loadSkillCFMap(),
      ]);
      const curDayoffMap = buildDayoffShiftMap(curDayoffRows);
      const curSkillFlat: SkillFlatRow[] = curSkillRows
        .map((row, i) => {
          const empId = cleanEmpId(row["Employee ID"] ?? row["Emp ID"] ?? "");
          const skill = String(row["Skill"] ?? "").trim();
          const level = Number(row["Level"] ?? row["level"]) || 0;
          return { id: `${i}-${empId}-${skill}`, empId, name: "", dept: "", jobSite: "", shift: "", shiftStart: "", dayoff: "", skill, level, origLevel: level };
        })
        .filter((r) => r.empId && r.skill);
      const currentCombined: CombinedEmployeeRow[] = curEmpRows.map((row) => {
        const empId = cleanEmpId(row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"]);
        if (!empId) return null;
        const ds = curDayoffMap.get(empId);
        const empSkills: Record<string, number> = {};
        for (const s of curSkillFlat.filter((r) => r.empId === empId)) empSkills[s.skill] = s.level;
        return {
          empId,
          firstName: String(row["First Name (Local)"] ?? "").trim(),
          lastName: String(row["Last Name (Local)"] ?? "").trim(),
          dept: String(row["หน่วยงาน"] ?? row["Name (Section)"] ?? "").trim(),
          position: String(row["Title (Position)"] ?? "").trim(),
          jobSite: resolveCombinedJobSite(ds?.section, curSkillCFMap.get(empId)),
          shift: ds?.shift ?? "",
          shiftStart: ds?.shiftStart ?? "",
          shiftEnd: ds?.shiftEnd ?? "",
          dayoff: ds?.dayoff ?? "",
          skills: empSkills,
        };
      }).filter(Boolean) as CombinedEmployeeRow[];

      const diff = computeEmployeeDiff(newCombinedRows, currentCombined, newManpowerRows);
      setDiffResult(diff);
      setShowDiffModal(true);
    } catch (e) {
      setError(e instanceof Error ? e.message : "อ่านไฟล์ไม่สำเร็จ");
    } finally {
      setIsSavingCombined(false);
    }
  }

  async function saveCombinedFromDiff() {
    if (!diffResult || !combinedFile) return;
    setIsSavingCombined(true);
    setError("");
    try {
      for (const fileType of ["employee_master", "dayoff_shift"] as const) {
        const fileId = crypto.randomUUID();
        const ext = getSafeFileExtension(combinedFile.name);
        const path = `${publicWorkspace}/masters/${fileType}/${fileId}${ext}`;
        const { error: uploadError } = await supabase.storage.from("workforce-inputs").upload(path, combinedFile, { upsert: true });
        if (uploadError) { setError(uploadError.message); return; }
        const { error: deactivateError } = await supabase.from("master_data_files").update({ is_active: false }).is("owner_id", null).eq("file_type", fileType);
        if (deactivateError) { setError(deactivateError.message); return; }
        const { error: insertError } = await supabase.from("master_data_files").insert({
          owner_id: null, file_type: fileType, file_path: path, original_filename: combinedFile.name, file_size_bytes: combinedFile.size, is_active: true,
        });
        if (insertError) { setError(insertError.message); return; }
      }
      const skillRows = diffResult.newRows.flatMap((r) =>
        Object.entries(r.skills)
          .filter(([, level]) => level > 0)
          .map(([skill, level]) => ({ "Employee ID": r.empId, Skill: skill, Level: level, "Can Do": 1 })),
      );
      const wsSkill = XLSX.utils.json_to_sheet(skillRows.length ? skillRows : [{ "Employee ID": "", Skill: "", Level: 0, "Can Do": 0 }]);
      const wbSkill = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wbSkill, wsSkill, "SkillMatrix");
      const skillBlob = new Blob([XLSX.write(wbSkill, { bookType: "xlsx", type: "array" })], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const skillPath = `${publicWorkspace}/masters/skill_matrix/${crypto.randomUUID()}.xlsx`;
      const { error: skillUpErr } = await supabase.storage.from("workforce-inputs").upload(skillPath, skillBlob, { upsert: true });
      if (skillUpErr) { setError(skillUpErr.message); return; }
      const { error: skillDeactErr } = await supabase.from("master_data_files").update({ is_active: false }).is("owner_id", null).eq("file_type", "skill_matrix");
      if (skillDeactErr) { setError(skillDeactErr.message); return; }
      const { error: skillInsErr } = await supabase.from("master_data_files").insert({ owner_id: null, file_type: "skill_matrix", file_path: skillPath, original_filename: "skill_matrix-from-combined.xlsx", file_size_bytes: skillBlob.size, is_active: true });
      if (skillInsErr) { setError(skillInsErr.message); return; }

      if (diffResult.newManpowerRows.length > 0) {
        const wsMp = XLSX.utils.json_to_sheet(diffResult.newManpowerRows);
        const wbMp = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wbMp, wsMp, "ManpowerPlan");
        const mpBlob = new Blob([XLSX.write(wbMp, { bookType: "xlsx", type: "array" })], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const mpPath = `${publicWorkspace}/masters/manpower_plan/${crypto.randomUUID()}.xlsx`;
        const { error: mpUpErr } = await supabase.storage.from("workforce-inputs").upload(mpPath, mpBlob, { upsert: true });
        if (mpUpErr) { setError(mpUpErr.message); return; }
        const { error: mpDeactErr } = await supabase.from("master_data_files").update({ is_active: false }).is("owner_id", null).eq("file_type", "manpower_plan");
        if (mpDeactErr) { setError(mpDeactErr.message); return; }
        const { error: mpInsErr } = await supabase.from("master_data_files").insert({ owner_id: null, file_type: "manpower_plan", file_path: mpPath, original_filename: "manpower_plan-from-combined.xlsx", file_size_bytes: mpBlob.size, is_active: true });
        if (mpInsErr) { setError(mpInsErr.message); return; }
      }

      const mpNote = diffResult.newManpowerRows.length === 0 ? " (Manpower Plan: ไม่มีข้อมูลใน Sheet2 — ข้ามการอัพเดท)" : "";
      setShowDiffModal(false);
      setDiffResult(null);
      setCombinedFile(null);
      setMessage(`บันทึกไฟล์รวมเรียบร้อย — พนักงาน ${diffResult.newRows.length} คน${mpNote}`);
      await onMastersSaved();
    } catch (e) {
      setError(e instanceof Error ? e.message : "บันทึกไม่สำเร็จ");
    } finally {
      setIsSavingCombined(false);
    }
  }

  return (
    <section className="md-page">

      {masterSubTab === "holidays" ? (
        <HolidayMasterPage onHolidaysChanged={onHolidaysChanged} />
      ) : null}

      {masterSubTab === "public_holidays" ? (
        <PublicHolidayPage onHolidaysChanged={onHolidaysChanged} />
      ) : null}

      {showDiffModal && diffResult && (
        <DiffPreviewModal
          diff={diffResult}
          isSaving={isSavingCombined}
          onCancel={() => { setShowDiffModal(false); setDiffResult(null); }}
          onConfirm={() => void saveCombinedFromDiff()}
        />
      )}

      {masterSubTab === "files" ? (<>

      <div className="md-two-col">

      {/* Combined template section */}
      <div className="master-card panel">
        <div className="master-card-header">
          <div className="combined-upload-info">
            <FileSpreadsheet size={18} />
            <div>
              <strong>ไฟล์รวมพนักงาน</strong>
              <span>Export ข้อมูลปัจจุบัน แก้ใน Excel แล้วอัพโหลดกลับ — ระบบจะแสดง diff ก่อน Save</span>
            </div>
          </div>
        </div>
        <label className={`ts-dropzone compact-dropzone ${combinedFile ? "has-file" : ""}`}>
          <UploadCloud size={24} />
          {combinedFile ? (
            <>
              <strong>{combinedFile.name}</strong>
              <span>{(combinedFile.size / 1024).toFixed(0)} KB · คลิกเพื่อเปลี่ยนไฟล์</span>
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
            onChange={(e) => setCombinedFile(e.target.files?.[0] ?? null)}
          />
        </label>
        <div className="master-card-history">
          {(() => {
            const files = masterFileHistory.filter((f) => f.file_type === "employee_master");
            const totalBytes = files.reduce((acc, f) => acc + (f.file_size_bytes ?? 0), 0);
            const totalMB = Math.round(totalBytes / (1024 * 1024));
            const pct = Math.min(100, Math.round((totalMB / STORAGE_LIMIT_MB) * 100));
            return (
              <div className="ts-history-header">
                <h4>ประวัติการอัปโหลด</h4>
                <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                  {totalBytes > 0 && (
                    <>
                      <div className="storage-bar-track" style={{ flex: "none", width: "90px" }}>
                        <div className={`storage-bar-fill${pct >= 80 ? " danger" : pct >= 60 ? " warn" : ""}`} style={{ width: `${Math.max(pct, 0.5)}%` }} />
                      </div>
                      <span className="storage-bar-label">~{totalMB} MB / {STORAGE_LIMIT_MB} MB ({pct}%)</span>
                    </>
                  )}
                  <span className="table-count">{files.length} ไฟล์</span>
                </div>
              </div>
            );
          })()}
          <div className="ts-history-list">
            {masterFileHistory.filter((f) => f.file_type === "employee_master").length === 0 ? (
              <p className="empty-copy" style={{ padding: "8px 0" }}>ยังไม่มีประวัติ</p>
            ) : null}
            {masterFileHistory
              .filter((f) => f.file_type === "employee_master")
              .map((file) => {
                const dateText = new Date(file.created_at).toLocaleString("th-TH", {
                  day: "numeric", month: "numeric", year: "2-digit",
                  hour: "2-digit", minute: "2-digit", hour12: false,
                });
                return (
                  <div className="ts-history-row" key={file.id}>
                    <div className="ts-history-info">
                      <div className="ts-history-name-row">
                        <strong>{file.original_filename ?? "ไฟล์"}</strong>
                        {file.is_active ? <span className="status-pill uploaded">Active</span> : null}
                      </div>
                      <span>{dateText}{file.file_size_bytes != null ? ` · ${(file.file_size_bytes / 1024).toFixed(0)} KB` : ""}</span>
                    </div>
                    <div className="ts-history-actions">
                      <button className="icon-button" onClick={() => downloadMasterFile(file.file_path, file.original_filename ?? "download.xlsx")} title="ดาวน์โหลด" type="button">
                        <Download size={15} />
                      </button>
                      <button className="icon-button danger" onClick={() => void onDeleteMasterFile(file)} title="ลบ" type="button">
                        <X size={15} />
                      </button>
                    </div>
                  </div>
                );
              })}
          </div>
        </div>
        <div className="master-card-actions">
          <button className="secondary-button small" onClick={downloadCombinedTemplate} type="button">
            <Download size={14} />
            เทมเพลตเปล่า
          </button>
          <button
            className="secondary-button small"
            disabled={isExporting || !activeMasterMap.employee_master}
            onClick={() => void exportCombinedMaster()}
            type="button"
          >
            <Download size={14} />
            {isExporting ? "กำลัง Export..." : "Export ข้อมูลปัจจุบัน"}
          </button>
          <button
            className="primary-button small"
            disabled={!combinedFile || isSavingCombined}
            onClick={() => void previewCombinedUpload()}
            type="button"
          >
            <UploadCloud size={14} />
            {isSavingCombined ? "กำลังอ่านไฟล์..." : "ตรวจสอบก่อนอัพโหลด"}
          </button>
        </div>
      </div>

      {(() => {
        const pendingFile = masterUploads.manpower_plan;
        const fileHistory = masterFileHistory.filter((f) => f.file_type === "manpower_plan");
        return (
          <div className="master-card master-card--orange panel">
            <div className="master-card-header">
              <div className="combined-upload-info" style={{ color: "#d97706" }}>
                <BarChart3 size={18} />
                <div>
                  <strong>Manpower Plan</strong>
                  <span>อัพเดทกะและเวลาเข้างานของแต่ละหน่วยงาน — ไม่กระทบรายชื่อพนักงาน</span>
                </div>
              </div>
            </div>
            <label className={`ts-dropzone compact-dropzone ${pendingFile ? "has-file" : ""}`}>
              <UploadCloud size={24} />
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
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={(e) => setMasterUploads((cur) => ({ ...cur, manpower_plan: e.target.files?.[0] ?? null }))}
              />
            </label>
            <div className="master-card-history">
              {(() => {
                const totalBytes = fileHistory.reduce((acc, f) => acc + (f.file_size_bytes ?? 0), 0);
                const totalMB = Math.round(totalBytes / (1024 * 1024));
                const pct = Math.min(100, Math.round((totalMB / STORAGE_LIMIT_MB) * 100));
                return (
                  <div className="ts-history-header">
                    <h4>ประวัติการอัปโหลด</h4>
                    <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                      {totalBytes > 0 && (
                        <>
                          <div className="storage-bar-track" style={{ flex: "none", width: "90px" }}>
                            <div className={`storage-bar-fill${pct >= 80 ? " danger" : pct >= 60 ? " warn" : ""}`} style={{ width: `${Math.max(pct, 0.5)}%` }} />
                          </div>
                          <span className="storage-bar-label">~{totalMB} MB / {STORAGE_LIMIT_MB} MB ({pct}%)</span>
                        </>
                      )}
                      <span className="table-count">{fileHistory.length} ไฟล์</span>
                    </div>
                  </div>
                );
              })()}
              <div className="ts-history-list">
                {fileHistory.length === 0 ? (
                  <p className="empty-copy" style={{ padding: "8px 0" }}>ยังไม่มีประวัติ</p>
                ) : null}
                {fileHistory.map((file) => {
                  const dateText = new Date(file.created_at).toLocaleString("th-TH", {
                    day: "numeric", month: "numeric", year: "2-digit",
                    hour: "2-digit", minute: "2-digit", hour12: false,
                  });
                  return (
                    <div className="ts-history-row" key={file.id}>
                      <div className="ts-history-info">
                        <div className="ts-history-name-row">
                          <strong>{file.original_filename ?? "ไฟล์"}</strong>
                          {file.is_active ? <span className="status-pill uploaded">Active</span> : null}
                        </div>
                        <span>{dateText}{file.file_size_bytes != null ? ` · ${(file.file_size_bytes / 1024).toFixed(0)} KB` : ""}</span>
                      </div>
                      <div className="ts-history-actions">
                        <button className="icon-button" onClick={() => downloadMasterFile(file.file_path, file.original_filename ?? "download.xlsx")} title="ดาวน์โหลด" type="button">
                          <Download size={15} />
                        </button>
                        <button className="icon-button danger" onClick={() => void onDeleteMasterFile(file)} title="ลบ" type="button">
                          <X size={15} />
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
            <div className="master-card-actions">
              <button
                className="primary-button small"
                disabled={!pendingFile || isSavingMasters}
                onClick={saveMasterFiles}
                type="button"
              >
                <UploadCloud size={14} />
                {isSavingMasters ? "Saving..." : "Save Manpower Plan"}
              </button>
            </div>
          </div>
        );
      })()}

      </div>{/* md-two-col */}

      </>) : null}

      {masterSubTab === "manpower" ? (
        <ManpowerEditor
          activeFile={activeMasterMap.manpower_plan}
          saveManpowerRows={saveManpowerRows}
        />
      ) : null}

      {masterSubTab === "dayoff_shift" ? (
        <DayoffShiftEditor
          activeFile={activeMasterMap.dayoff_shift}
          employeeMasterFile={activeMasterMap.employee_master}
          manpowerFile={activeMasterMap.manpower_plan}
          saveDayoffShiftRows={saveDayoffShiftRows}
          saveManpowerRows={saveManpowerRows}
        />
      ) : null}
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
    const existing = holidays.find((h) => h.date === newDate);
    if (existing) {
      alert(`วันที่ ${newDate} มีอยู่แล้วเป็น "${existing.name}" (${existing.type})\nหากต้องการแก้ไข กรุณาลบรายการเดิมก่อน`);
      return;
    }
    setSaving(true);
    const { error } = await supabase
      .from("holidays")
      .insert({ date: newDate, name: newName.trim(), type: newType });
    if (error) {
      alert(`เพิ่มไม่สำเร็จ: ${error.message}`);
    } else {
      setNewDate("");
      setNewName("");
      await loadHolidays();
    }
    setSaving(false);
  }

  async function deleteHoliday(id: string) {
    const { error } = await supabase.from("holidays").delete().eq("id", id);
    if (error) { alert(`ลบไม่สำเร็จ: ${error.message}`); return; }
    await loadHolidays();
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

function PublicHolidayPage({
  onHolidaysChanged,
}: {
  onHolidaysChanged: (dates: Set<string>) => void;
}) {
  const [holidays, setHolidays] = useState<HolidayRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [newDate, setNewDate] = useState("");
  const [newName, setNewName] = useState("");
  const [newType, setNewType] = useState<"public_holiday" | "company_holiday">("public_holiday");
  const [yearFilter, setYearFilter] = useState(String(new Date().getFullYear()));
  const seededRef = useRef(false);

  useEffect(() => {
    void loadHolidays();
  }, []);

  async function loadHolidays() {
    setLoading(true);

    if (!seededRef.current) {
      seededRef.current = true;
      // fetch ALL dates to avoid overwriting Buddhist holy days
      const { data: allDates } = await supabase.from("holidays").select("date");
      const allExistingDates = new Set((allDates ?? []).map((r: { date: string }) => r.date));

      const toInsert = Object.values(cpfPublicHolidaysByYear)
        .flat()
        .filter((e) => !allExistingDates.has(e.date))
        .map((e) => ({ date: e.date, name: e.name, type: "public_holiday" as const }));

      if (toInsert.length > 0) {
        const { error: insertError } = await supabase.from("holidays").insert(toInsert);
        if (insertError) console.error("[PublicHolidayPage] auto-seed failed:", insertError.message);
      }
    }

    // fetch public/company holidays + Buddhist holy days that overlap with CPF dates
    const cpfDatesList = Object.values(cpfPublicHolidaysByYear).flat().map((e) => e.date);
    const { data } = await supabase
      .from("holidays")
      .select("*")
      .or(`type.in.(public_holiday,company_holiday),date.in.(${cpfDatesList.join(",")})`)
      .order("date");

    if (data) {
      setHolidays(data as HolidayRow[]);
      const { data: all } = await supabase.from("holidays").select("date");
      if (all) onHolidaysChanged(new Set(all.map((r: { date: string }) => r.date)));
    }
    setLoading(false);
  }

  async function addHoliday() {
    if (!newDate || !newName.trim()) return;
    const existing = holidays.find((h) => h.date === newDate);
    if (existing) {
      alert(`วันที่ ${newDate} มีอยู่แล้วเป็น "${existing.name}" (${existing.type})\nหากต้องการแก้ไข กรุณาลบรายการเดิมก่อน`);
      return;
    }
    setSaving(true);
    const { error } = await supabase
      .from("holidays")
      .insert({ date: newDate, name: newName.trim(), type: newType });
    if (error) {
      alert(`เพิ่มไม่สำเร็จ: ${error.message}`);
    } else {
      setNewDate("");
      setNewName("");
      await loadHolidays();
    }
    setSaving(false);
  }

  async function deleteHoliday(id: string) {
    const { error } = await supabase.from("holidays").delete().eq("id", id);
    if (error) { alert(`ลบไม่สำเร็จ: ${error.message}`); return; }
    await loadHolidays();
  }

  const cpfDateNameMap = new Map<string, string>(
    Object.values(cpfPublicHolidaysByYear).flat().map((e) => [e.date, e.name])
  );
  const cpfDateSet = new Set(cpfDateNameMap.keys());

  const existingYears = Array.from(new Set(holidays.map((h) => h.date.substring(0, 4)))).sort();
  const allYears = Array.from(new Set([...existingYears, yearFilter])).sort();
  // show only CPF dates (exclude Buddhist-only rows that don't match CPF)
  const filtered = holidays
    .filter((h) => h.date.startsWith(yearFilter) && (h.type !== "buddhist_holy_day" || cpfDateSet.has(h.date)));

  const typeLabel: Record<string, string> = {
    public_holiday: "วันหยุดราชการ",
    company_holiday: "วันหยุดบริษัท",
    buddhist_holy_day: "วันหยุดราชการ",
  };
  const typeBadgeClass: Record<string, string> = {
    public_holiday: "holiday-badge-public",
    company_holiday: "holiday-badge-company",
    buddhist_holy_day: "holiday-badge-public",
  };

  return (
    <section className="panel holiday-master-panel">
      <div className="holiday-master-header">
        <div className="holiday-master-title">
          <CalendarDays size={20} />
          <h2>วันหยุดประจำปี</h2>
          <span className="holiday-total-badge">{holidays.length} วัน</span>
        </div>
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
            onChange={(e) => setNewType(e.target.value as "public_holiday" | "company_holiday")}
          >
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
                <th>หมายเหตุ</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((h) => {
                const d = new Date(h.date + "T00:00:00");
                const isBuddhist = isBuddhistHolyDay(d);
                const displayName = h.type === "buddhist_holy_day"
                  ? (cpfDateNameMap.get(h.date) ?? h.name)
                  : h.name;
                return (
                  <tr key={h.id}>
                    <td className="holiday-date-cell">
                      {d.toLocaleDateString("th-TH", { day: "numeric", month: "long", year: "numeric" })}
                    </td>
                    <td>
                      {["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"][d.getDay()]}
                    </td>
                    <td>{displayName}</td>
                    <td>
                      <span className={`holiday-type-badge ${typeBadgeClass[h.type]}`}>
                        {typeLabel[h.type]}
                      </span>
                    </td>
                    <td>
                      {isBuddhist && (
                        <span className="holiday-type-badge holiday-badge-buddhist">วันพระ</span>
                      )}
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
                  <td colSpan={6} className="holiday-empty-row">
                    ไม่มีวันหยุดสำหรับปี {yearFilter} — เพิ่มเองด้านบน
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

function ManpowerEditor({
  activeFile,
  saveManpowerRows,
}: {
  activeFile?: MasterFile;
  saveManpowerRows: (rows: ManpowerEditorRow[]) => Promise<void>;
}) {
  const [rows, setRows] = useState<ManpowerEditorRow[]>([]);
  const [originalRows, setOriginalRows] = useState<ManpowerEditorRow[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [newDept, setNewDept] = useState("");
  const [newJobSite, setNewJobSite] = useState("");
  const [newShift, setNewShift] = useState("");
  const [newShiftStart, setNewShiftStart] = useState("");
  const [newShiftEnd, setNewShiftEnd] = useState("");
  const [deptFilter, setDeptFilter] = useState("all");
  const [jobSiteFilter, setJobSiteFilter] = useState("all");
  const [shiftFilter, setShiftFilter] = useState("all");
  const [sort, setSort_] = useState<SortState>(null);
  const setSort = setSort_ as (sort: SortState) => void;

  useEffect(() => {
    if (!activeFile?.file_path) {
      setRows([]);
      setOriginalRows([]);
      return;
    }
    let isMounted = true;
    setIsLoading(true);
    downloadSheetRows(activeFile.file_path)
      .then((sheetRows) => {
        if (!isMounted) return;
        const parsed = sheetRows
          .map((row, i) => ({
            id: `mp-${i}`,
            dept: String(row["หน่วยงาน"] ?? row["dept"] ?? "").trim(),
            jobSite: String(row["หน่วยงานย่อย"] ?? row["หน้างาน"] ?? row["job_site"] ?? "").trim(),
            shift: normalizeShiftLabel(row["กะ"] ?? row["shift"] ?? row["อยู่กะไหน"]),
            shiftStart: normalizeTimeText(row["เวลาเข้า"] ?? row["shift_start"] ?? row["เวลาเข้างาน"]),
            shiftEnd: normalizeTimeText(row["เวลาออก"] ?? row["เวลาออกงาน"] ?? row["shift_end"]),
          }))
          .filter((r) => r.dept || r.shift);
        setRows(parsed);
        setOriginalRows(parsed);
      })
      .catch(() => {
        if (!isMounted) return;
        setRows([]);
        setOriginalRows([]);
      })
      .finally(() => {
        if (isMounted) setIsLoading(false);
      });
    return () => { isMounted = false; };
  }, [activeFile?.file_path]);

  const isDirty = JSON.stringify(rows) !== JSON.stringify(originalRows);

  function updateRow(id: string, field: "shiftStart" | "shiftEnd", value: string) {
    setRows((current) =>
      current.map((row) => {
        if (row.id !== id) return row;
        const next = { ...row, [field]: value };
        if (field === "shiftStart" && !row.shiftEnd) next.shiftEnd = addHoursToTime(value, 9);
        return next;
      }),
    );
  }

  function rowKey(row: ManpowerEditorRow) {
    return `${row.dept}||${row.jobSite}||${normalizeShiftKey(row.shift)}`;
  }

  function addRow() {
    if (!newDept.trim() || !newShift.trim()) return;
    if (newShift.trim() === "ผู้จัดการ") {
      alert(
        `"ผู้จัดการ" เป็นคำสงวนที่ระบบใช้หมายถึง "ไม่มีกะตายตัว" (ไม่ล็อคเวลา)\n` +
        `ตั้งเป็นชื่อกะจริงไม่ได้ — เวลาที่กรอกไว้จะไม่ถูกใช้งานเลย\nกรุณาใช้ชื่อกะอื่น`,
      );
      return;
    }
    const shiftStart = newShiftStart;
    const shiftEnd = newShiftEnd || (shiftStart ? addHoursToTime(shiftStart, 9) : "");
    const candidate: ManpowerEditorRow = {
      id: `mp-new-${Date.now()}`,
      dept: newDept.trim(),
      jobSite: newJobSite.trim(),
      shift: newShift.trim(),
      shiftStart,
      shiftEnd,
    };
    const exists = rows.some((r) => r.dept && r.shift && rowKey(r) === rowKey(candidate));
    if (exists) {
      alert(`หน่วยงาน "${candidate.dept}"${candidate.jobSite ? ` (${candidate.jobSite})` : ""} กะ "${candidate.shift}" มีอยู่แล้ว\nหากต้องการแก้ไข กรุณาแก้ที่แถวเดิมในตาราง`);
      return;
    }
    setRows((current) => [...current, candidate]);
    setNewDept("");
    setNewJobSite("");
    setNewShift("");
    setNewShiftStart("");
    setNewShiftEnd("");
  }

  function deleteRow(id: string) {
    setRows((current) => current.filter((r) => r.id !== id));
  }

  const duplicateKeys = (() => {
    const seen = new Set<string>();
    const dupes = new Set<string>();
    for (const row of rows) {
      if (!row.dept || !row.shift) continue;
      const key = rowKey(row);
      if (seen.has(key)) dupes.add(key);
      seen.add(key);
    }
    return dupes;
  })();

  const modifiedIds = new Set(
    rows
      .filter((row) => {
        const orig = originalRows.find((r) => r.id === row.id);
        return orig && (orig.shiftStart !== row.shiftStart || orig.shiftEnd !== row.shiftEnd);
      })
      .map((r) => r.id),
  );

  async function handleSave() {
    if (duplicateKeys.size > 0) {
      alert("มีหน่วยงาน + กะ ซ้ำกัน กรุณาแก้ไขก่อนบันทึก");
      return;
    }
    setIsSaving(true);
    try {
      await saveManpowerRows(rows);
      setOriginalRows(rows);
    } catch {
      // error already set by saveManpowerRows
    } finally {
      setIsSaving(false);
    }
  }

  const deptFilterOptions = Array.from(new Set(rows.map((r) => r.dept).filter(Boolean))).sort();
  const jobSiteFilterOptions = Array.from(
    new Set(rows.filter((r) => deptFilter === "all" || r.dept === deptFilter).map((r) => r.jobSite).filter(Boolean)),
  ).sort();
  const shiftFilterOptions = Array.from(new Set(rows.map((r) => r.shift).filter(Boolean))).sort();

  const filteredRows = rows.filter((row) => {
    if (deptFilter !== "all" && row.dept !== deptFilter) return false;
    if (jobSiteFilter !== "all" && row.jobSite !== jobSiteFilter) return false;
    if (shiftFilter !== "all" && row.shift !== shiftFilter) return false;
    return true;
  });

  const sortedRows = sort
    ? [...filteredRows].sort((a, b) => {
        const aValue = (a as unknown as Record<string, string>)[sort.key] ?? "";
        const bValue = (b as unknown as Record<string, string>)[sort.key] ?? "";
        const comparison = aValue.localeCompare(bValue, "th", { numeric: true, sensitivity: "base" });
        return sort.direction === "asc" ? comparison : -comparison;
      })
    : filteredRows;

  function handleExport() {
    const data = sortedRows.map((row) => ({
      "หน่วยงาน": row.dept,
      "หน่วยงานย่อย": row.jobSite,
      "กะ": row.shift,
      "เวลาเข้า": row.shiftStart,
      "เวลาออก": row.shiftEnd,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Manpower Plan");
    XLSX.writeFile(wb, "manpower-plan-export.xlsx");
  }

  return (
    <section className="panel holiday-master-panel">
      <div className="ot-detail-hdr">
        <div className="ot-detail-filters no-wrap">
          <div className="holiday-master-title" style={{ flexShrink: 0 }}>
            <BarChart3 size={20} />
            <h2 style={{ whiteSpace: "nowrap" }}>Manpower Plan</h2>
            {modifiedIds.size > 0 && <span className="modified-badge">{modifiedIds.size} แก้ไข</span>}
          </div>
          <select
            aria-label="หน่วยงาน"
            value={deptFilter}
            onChange={(e) => { setDeptFilter(e.target.value); setJobSiteFilter("all"); }}
          >
            <option value="all">ทุกหน่วยงาน</option>
            {deptFilterOptions.map((d) => <option key={d} value={d}>{d}</option>)}
          </select>
          <select
            aria-label="หน่วยงานย่อย"
            value={jobSiteFilter}
            onChange={(e) => setJobSiteFilter(e.target.value)}
            style={{ maxWidth: 260 }}
          >
            <option value="all">ทุกหน่วยงานย่อย</option>
            {jobSiteFilterOptions.map((j) => <option key={j} value={j}>{j}</option>)}
          </select>
          <select aria-label="กะ" value={shiftFilter} onChange={(e) => setShiftFilter(e.target.value)}>
            <option value="all">ทุกกะ</option>
            {shiftFilterOptions.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
          <button
            className="ghost-button"
            onClick={() => { setDeptFilter("all"); setJobSiteFilter("all"); setShiftFilter("all"); }}
            type="button"
          >
            Clear
          </button>
          <button type="button" className="primary-button small" onClick={handleExport} disabled={filteredRows.length === 0}>
            <Download size={14} />
            Export Excel
          </button>
          <button className="primary-button" style={{ flexShrink: 0 }} disabled={!rows.length || isSaving} onClick={handleSave} type="button">
            <UploadCloud size={17} />
            {isSaving ? "Saving..." : `Save${isDirty ? " (มีการแก้ไข)" : ""}`}
          </button>
        </div>
      </div>
      <p style={{ margin: "-12px 0 0", fontSize: 13, color: "var(--muted)" }}>
        กำหนดเวลาเข้า-ออกของแต่ละหน่วยงาน + กะ — หน้า Shift & Dayoff จะดึงเวลานี้ไปใช้อัตโนมัติ
      </p>

      <div className="holiday-add-form" style={{ marginTop: -12 }}>
        <div className="holiday-add-form-fields">
          <input
            type="text"
            className="holiday-input"
            placeholder="หน่วยงาน เช่น งานเครื่องใน"
            value={newDept}
            onChange={(e) => setNewDept(e.target.value)}
            style={{ flex: 2, minWidth: 160 }}
          />
          <input
            type="text"
            className="holiday-input"
            placeholder="หน่วยงานย่อย (ว่าง = ทั้งหน่วยงาน)"
            value={newJobSite}
            onChange={(e) => setNewJobSite(e.target.value)}
            style={{ flex: 2, minWidth: 160 }}
          />
          <input
            type="text"
            className="holiday-input"
            placeholder="กะ เช่น กะ1"
            value={newShift}
            onChange={(e) => setNewShift(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && addRow()}
            style={{ flex: 1, minWidth: 90 }}
          />
          <TimeInput24
            value={newShiftStart}
            onChange={setNewShiftStart}
            title="เวลาเข้า"
            className="time24-btn holiday-input"
          />
          <TimeInput24
            value={newShiftEnd}
            onChange={setNewShiftEnd}
            title="เวลาออก"
            className="time24-btn holiday-input"
          />
          <button
            className="primary-button"
            onClick={addRow}
            disabled={!newDept.trim() || !newShift.trim()}
            type="button"
          >
            + เพิ่มหน่วยงาน/กะ
          </button>
        </div>
      </div>

      <div className="dayoff-editor-table" style={{ marginTop: -12 }}>
        <table className="table">
          <thead>
            <tr>
              <th><SortButton columnKey="dept" setSort={setSort} sort={sort} defaultDirection="desc">หน่วยงาน</SortButton></th>
              <th><SortButton columnKey="jobSite" setSort={setSort} sort={sort} defaultDirection="desc">หน่วยงานย่อย</SortButton></th>
              <th><SortButton columnKey="shift" setSort={setSort} sort={sort} defaultDirection="desc">กะ</SortButton></th>
              <th><SortButton columnKey="shiftStart" setSort={setSort} sort={sort} defaultDirection="desc">เวลาเข้า</SortButton></th>
              <th><SortButton columnKey="shiftEnd" setSort={setSort} sort={sort} defaultDirection="desc">เวลาออก</SortButton></th>
              <th style={{ width: 40 }} />
            </tr>
          </thead>
          <tbody>
            {sortedRows.map((row) => (
              <tr
                key={row.id}
                className={(row.dept && row.shift && duplicateKeys.has(rowKey(row))) || modifiedIds.has(row.id) ? "row-modified" : ""}
              >
                <td>{row.dept || "—"}</td>
                <td>{row.jobSite || <span style={{ color: "#94a3b8" }}>ทั้งหน่วยงาน</span>}</td>
                <td>{row.shift || "—"}</td>
                <td>
                  <TimeInput24
                    value={row.shiftStart}
                    onChange={(v) => updateRow(row.id, "shiftStart", v)}
                    className="time24-btn shift-time-input"
                  />
                </td>
                <td className="shift-end-cell">
                  <TimeInput24
                    value={row.shiftEnd}
                    onChange={(v) => updateRow(row.id, "shiftEnd", v)}
                    className="time24-btn shift-time-input"
                  />
                </td>
                <td>
                  <button className="icon-button danger" onClick={() => deleteRow(row.id)} title="ลบแถว" type="button">
                    <Trash2 size={15} />
                  </button>
                </td>
              </tr>
            ))}
            {activeFile && isLoading ? (
              <tr><td colSpan={6}>Loading Manpower Plan...</td></tr>
            ) : null}
            {!isLoading && rows.length === 0 ? (
              <tr><td colSpan={6}>ยังไม่มีข้อมูล — กรอกฟอร์มด้านบนแล้วกด &quot;+ เพิ่มหน่วยงาน/กะ&quot; เพื่อเริ่มตั้งค่า</td></tr>
            ) : null}
            {!isLoading && rows.length > 0 && filteredRows.length === 0 ? (
              <tr><td colSpan={6}>ไม่พบข้อมูลที่ค้นหา</td></tr>
            ) : null}
          </tbody>
        </table>
      </div>
      {duplicateKeys.size > 0 && (
        <p style={{ fontSize: 12, color: "#ef4444", marginTop: 6 }}>
          พบหน่วยงาน + กะ ซ้ำกัน — ระบบจะใช้แถวแรกที่เจอเท่านั้น กรุณาลบแถวที่ซ้ำก่อนบันทึก
        </p>
      )}
    </section>
  );
}

function DayoffShiftEditor({
  activeFile,
  employeeMasterFile,
  manpowerFile,
  saveDayoffShiftRows,
  saveManpowerRows,
}: {
  activeFile?: MasterFile;
  employeeMasterFile?: MasterFile;
  manpowerFile?: MasterFile;
  saveDayoffShiftRows: (rows: DayoffShiftEditorRow[]) => Promise<void>;
  saveManpowerRows: (rows: ManpowerEditorRow[]) => Promise<void>;
}) {
  const [rows, setRows] = useState<DayoffShiftEditorRow[]>([]);
  const [originalRows, setOriginalRows] = useState<DayoffShiftEditorRow[]>([]);
  const [jobSiteOptions, setJobSiteOptions] = useState<string[]>([]);
  const [manpowerRows, setManpowerRows] = useState<Array<{ dept: string; jobSite: string; shift: string; shiftStart: string; shiftEnd: string }>>([]);
  const [query, setQuery] = useState("");
  const [selectedDept, setSelectedDept] = useState("all");
  const [selectedJobSite, setSelectedJobSite] = useState("all");
  const [selectedDayoff, setSelectedDayoff] = useState("all");
  const [selectedShift, setSelectedShift] = useState("all");
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [bulkDayoff, setBulkDayoff] = useState("");
  const [bulkShift, setBulkShift] = useState("");
  const [bulkJobSite, setBulkJobSite] = useState("");
  const [bulkShiftStart, setBulkShiftStart] = useState("");
  const [bulkShiftEnd, setBulkShiftEnd] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  // เวลาที่แก้ไขในหน้านี้ยังไม่ถูกบันทึกลง Manpower Plan จริง — ต้องกด Save เพื่อบันทึกทั้งคู่
  const [manpowerDirty, setManpowerDirty] = useState(false);

  useEffect(() => {
    const path = manpowerFile?.file_path;
    if (!path) { setManpowerRows([]); setManpowerDirty(false); return; }
    let cancelled = false;
    downloadSheetRows(path)
      .then((sheetRows) => {
        if (cancelled) return;
        // ใช้ parser เดียวกับ buildReportData เพื่อไม่ให้แถว default ของหน่วยงาน (ไม่ระบุกะ) หายไปจากหน้านี้
        // ทั้งที่ Dashboard/OT จริงยังเห็นและใช้แถวนั้นอยู่
        setManpowerRows(parseManpowerRows(sheetRows));
        setManpowerDirty(false);
      })
      .catch(() => { if (!cancelled) setManpowerRows([]); });
    return () => { cancelled = true; };
  }, [manpowerFile?.file_path]);

  // ใช้ตัว lookup เดียวกับที่ buildReportData ใช้คำนวณ Dashboard/OT จริง เพื่อไม่ให้เวลาที่ "ล็อค" ในหน้านี้
  // กับเวลาที่ Dashboard คำนวณจริงเพี้ยนกันสำหรับพนักงานที่ยังไม่เคย save ผ่านหน้านี้
  const manpowerLookup = useMemo(() => buildManpowerLookup(manpowerRows), [manpowerRows]);

  function lookupManpower(dept: string, jobSite: string, shift: string) {
    return lookupManpowerTime(manpowerLookup, dept, jobSite, shift);
  }

  // เจอแถวใน Manpower ที่ dept+หน่วยงานย่อย+กะ+เวลา ตรงกันเป๊ะแล้วหรือยัง
  function findManpowerExact(dept: string, jobSite: string, shift: string, shiftStart: string, shiftEnd: string) {
    const shiftKey = normalizeShiftKey(shift);
    return manpowerRows.find(
      (r) => r.dept === dept && r.jobSite === jobSite && normalizeShiftKey(r.shift) === shiftKey
        && r.shiftStart === shiftStart && r.shiftEnd === shiftEnd,
    );
  }

  // dept+หน่วยงานย่อย+กะ นี้มีเวลาอื่น (คนละเวลากับที่กำลังจะตั้ง) บันทึกอยู่แล้วหรือไม่ — ถ้ามีต้องตั้งชื่อ
  // หน่วยงานย่อยใหม่กันชนกับของเดิม (เหมือนที่แก้ปัญหา QC มีหลายเวลาไปแล้วก่อนหน้านี้)
  function hasManpowerCollision(dept: string, jobSite: string, shift: string) {
    const shiftKey = normalizeShiftKey(shift);
    return manpowerRows.some((r) => r.dept === dept && r.jobSite === jobSite && normalizeShiftKey(r.shift) === shiftKey);
  }

  // ตัดส่วนต่อท้าย " (HH:MM-HH:MM)" ที่ระบบเติมเองตอนสร้างกะใหม่ออกก่อนเช็คซ้ำ — กันไม่ให้ต่อท้ายซ้อนกันหลายชั้น
  // เวลาผู้ใช้แก้เวลาเข้ากับเวลาออกทีละครั้ง (2 อีเวนต์แยกกัน) จากหน่วยงานย่อยที่เพิ่งถูกตั้งชื่อใหม่ไปแล้ว
  function stripTimeSuffix(jobSite: string): string {
    return jobSite.replace(/\s\(\d{2}:\d{2}-\d{2}:\d{2}\)$/, "");
  }

  // ตัดสินใจว่าเวลาใหม่ที่ตั้งควรใช้ชื่อหน่วยงานย่อยเดิม หรือต้องตั้งชื่อใหม่ต่อท้ายด้วยเวลา
  function resolveTimeChange(dept: string, jobSite: string, shift: string, shiftStart: string, shiftEnd: string) {
    const baseJobSite = stripTimeSuffix(jobSite);
    if (findManpowerExact(dept, baseJobSite, shift, shiftStart, shiftEnd)) {
      return { jobSite: baseJobSite, isNew: false };
    }
    const collision = hasManpowerCollision(dept, baseJobSite, shift);
    const finalJobSite = collision ? `${baseJobSite || "ทั้งหน่วยงาน"} (${shiftStart}-${shiftEnd})` : baseJobSite;
    return { jobSite: finalJobSite, isNew: true };
  }

  // เพิ่มแถว Manpower ใหม่ (ถ้ายังไม่มีแถวนี้เป๊ะๆ อยู่แล้ว) และตั้งสถานะว่ามีการแก้ไข Manpower รอบันทึก
  function addManpowerRowIfNeeded(dept: string, jobSite: string, shift: string, shiftStart: string, shiftEnd: string) {
    setManpowerRows((prev) => {
      const shiftKey = normalizeShiftKey(shift);
      const exists = prev.some(
        (r) => r.dept === dept && r.jobSite === jobSite && normalizeShiftKey(r.shift) === shiftKey
          && r.shiftStart === shiftStart && r.shiftEnd === shiftEnd,
      );
      if (exists) return prev;
      return [...prev, { dept, jobSite, shift, shiftStart, shiftEnd }];
    });
    setManpowerDirty(true);
  }

  const manpowerShiftsByKey = useMemo(() => {
    const map = new Map<string, string[]>();
    for (const r of manpowerRows) {
      if (!r.shift) continue; // แถว default ของหน่วยงาน (ไม่ระบุกะ) ไม่ใช่ตัวเลือกกะที่เลือกได้จริง
      const key = `${r.dept}||${r.jobSite}`;
      const list = map.get(key) ?? [];
      if (!list.includes(r.shift)) list.push(r.shift);
      map.set(key, list);
    }
    return map;
  }, [manpowerRows]);

  // กะทั้งหมดที่หน่วยงานนี้มีจริงใน Manpower (รวมทุกหน่วยงานย่อย) — ใช้เป็น fallback ก่อนจะหลุดไปใช้ลิสต์รวมทั้งบริษัท
  const manpowerShiftsByDept = useMemo(() => {
    const map = new Map<string, string[]>();
    for (const r of manpowerRows) {
      if (!r.shift) continue; // แถว default ของหน่วยงาน (ไม่ระบุกะ) ไม่ใช่ตัวเลือกกะที่เลือกได้จริง
      const list = map.get(r.dept) ?? [];
      if (!list.includes(r.shift)) list.push(r.shift);
      map.set(r.dept, list);
    }
    return map;
  }, [manpowerRows]);

  // ดึงหน่วยงานย่อยที่ตั้งไว้ในหน้า Manpower ของหน่วยงานนั้น มารวมกับตัวเลือกเดิม (จากไฟล์รวมพนักงาน)
  const manpowerJobSitesByDept = useMemo(() => {
    const map = new Map<string, string[]>();
    for (const r of manpowerRows) {
      if (!r.jobSite) continue;
      const list = map.get(r.dept) ?? [];
      if (!list.includes(r.jobSite)) list.push(r.jobSite);
      map.set(r.dept, list);
    }
    return map;
  }, [manpowerRows]);

  // หน่วยงานย่อยที่พนักงานคนอื่นในหน่วยงานเดียวกันใช้อยู่แล้ว (นอกจาก Manpower)
  const employeeJobSitesByDept = useMemo(() => {
    const map = new Map<string, Set<string>>();
    for (const r of rows) {
      if (!r.jobSite) continue;
      const set = map.get(r.dept) ?? new Set<string>();
      set.add(r.jobSite);
      map.set(r.dept, set);
    }
    return map;
  }, [rows]);

  // ค่าฐานของตัวเลือกหน่วยงานย่อยต่อหน่วยงาน คำนวณครั้งเดียวต่อหน่วยงานที่พบจริง (เดิม spread+sort ทุกแถว)
  // จำกัดขอบเขตแค่หน่วยงานย่อยที่เกี่ยวข้องกับหน่วยงานนั้นจริง (จาก Manpower + คนอื่นในหน่วยงานเดียวกัน)
  // แทนที่จะรวมลิสต์ทั้งบริษัท (~40 รายการ) ในทุกแถว — ลดจำนวน DOM option ต่อแถวลงมาก
  // ถ้าหน่วยงานนั้นไม่มีข้อมูลเลยทั้งสองแหล่ง ค่อย fallback เป็นลิสต์รวมทั้งบริษัท
  const jobSiteChoiceBaseByDept = useMemo(() => {
    const map = new Map<string, string[]>();
    for (const dept of new Set(rows.map((r) => r.dept))) {
      const fromMp = manpowerJobSitesByDept.get(dept) ?? [];
      const fromEmp = Array.from(employeeJobSitesByDept.get(dept) ?? []);
      const combined = Array.from(new Set([...fromMp, ...fromEmp])).sort();
      map.set(dept, combined.length > 0 ? combined : jobSiteOptions);
    }
    return map;
  }, [rows, manpowerJobSitesByDept, employeeJobSitesByDept, jobSiteOptions]);

  function jobSiteChoicesForDept(dept: string, current: string) {
    const merged = jobSiteChoiceBaseByDept.get(dept) ?? jobSiteOptions;
    return current && !merged.includes(current) ? [...merged, current] : merged;
  }

  // หน่วยงาน+กะ ที่ Manpower Plan กำหนดหน่วยงานย่อยไว้ "ค่าเดียวไม่ก้ำกึ่ง" เท่านั้น — ใช้แนะนำอัตโนมัติได้อย่างปลอดภัย
  const uniqueJobSiteByDeptShift = useMemo(() => {
    const seen = new Map<string, Set<string>>();
    for (const r of manpowerRows) {
      if (!r.jobSite) continue;
      const key = makeDeptShiftKey(r.dept, r.shift);
      const set = seen.get(key) ?? new Set<string>();
      set.add(r.jobSite);
      seen.set(key, set);
    }
    const map = new Map<string, string>();
    for (const [key, set] of seen) {
      if (set.size === 1) map.set(key, Array.from(set)[0]);
    }
    return map;
  }, [manpowerRows]);

  function computeJobSiteSuggestions() {
    return rows
      .map((row) => {
        if (!row.dept || !row.shift) return null;
        const suggested = uniqueJobSiteByDeptShift.get(makeDeptShiftKey(row.dept, row.shift));
        if (!suggested || suggested === row.jobSite) return null;
        return { row, suggested };
      })
      .filter((x): x is { row: DayoffShiftEditorRow; suggested: string } => x !== null);
  }

  function syncJobSitesFromManpower() {
    const suggestions = computeJobSiteSuggestions();
    if (suggestions.length === 0) {
      alert("ไม่พบรายการที่ต้องปรับ — หน่วยงานย่อยตรงกับ Manpower Plan อยู่แล้ว หรือ Manpower Plan ยังไม่ได้ระบุหน่วยงานย่อยแบบไม่ก้ำกึ่งสำหรับหน่วยงาน+กะนั้นๆ");
      return;
    }
    const preview = suggestions
      .slice(0, 15)
      .map((s) => `${s.row.empId} ${s.row.name}: "${s.row.jobSite || "-"}" → "${s.suggested}"`)
      .join("\n");
    const more = suggestions.length > 15 ? `\n...และอีก ${suggestions.length - 15} คน` : "";
    const ok = window.confirm(
      `พบ ${suggestions.length} คนที่หน่วยงานย่อยไม่ตรงกับ Manpower Plan (มีค่าที่ควรจะเป็นชัดเจนไม่ก้ำกึ่ง)\nต้องการปรับหน่วยงานย่อยให้ตรงกันหรือไม่?\n\n${preview}${more}\n\n(ยังไม่บันทึก — ต้องกด Save อีกครั้งหลังตรวจสอบ)`,
    );
    if (!ok) return;
    const suggestionMap = new Map(suggestions.map((s) => [s.row.id, s.suggested]));
    setRows((current) =>
      current.map((row) => {
        const suggested = suggestionMap.get(row.id);
        if (!suggested) return row;
        const raw = setRowCol(row.raw, suggested, "หน่วยงานย่อย/Skill", "หน้างาน", "job_site", "Job Site");
        return { ...row, jobSite: suggested, raw };
      }),
    );
  }

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
  // ต้อง useMemo (ไม่ใช่ IIFE เฉยๆ) เพราะ shiftChoiceBaseByRowKey ด้านล่างเอาไปเป็น dependency —
  // ถ้าเป็น array อ้างอิงใหม่ทุก render จะทำให้ memo นั้นคำนวณใหม่ทุกครั้งโดยไม่จำเป็น (พิมพ์ในกล่องค้นหาก็รีคำนวณ)
  const shiftOptions = useMemo(() => {
    // dedupe by normalized shift key (กะ1 vs กะ 1) — prefer employees' own stored label so filtering matches exactly
    const byKey = new Map<string, string>();
    for (const r of manpowerRows) { if (r.shift) byKey.set(normalizeShiftKey(r.shift), r.shift); }
    for (const r of rows) { if (r.shift) byKey.set(normalizeShiftKey(r.shift), r.shift); }
    return Array.from(byKey.values()).sort();
  }, [manpowerRows, rows]);

  // ค่าฐานของตัวเลือกกะต่อ (หน่วยงาน, หน่วยงานย่อย) หนึ่งคู่ คำนวณครั้งเดียวต่อคู่ที่พบจริง
  // ไม่ใช่คำนวณใหม่ทุกแถวพนักงาน (439 คน) — จุดที่ทำให้ตารางช้าตอน render/clear filter
  const shiftChoiceBaseByRowKey = useMemo(() => {
    const map = new Map<string, string[]>();
    for (const row of rows) {
      const key = `${row.dept}||${row.jobSite}`;
      if (map.has(key)) continue;
      const specific = manpowerShiftsByKey.get(key) ?? [];
      const deptWide = manpowerShiftsByKey.get(`${row.dept}||`) ?? [];
      const deptAny = manpowerShiftsByDept.get(row.dept) ?? [];
      const fromMp = specific.length > 0 ? specific : deptWide.length > 0 ? deptWide : deptAny;
      map.set(key, fromMp.length > 0 ? fromMp : shiftOptions);
    }
    return map;
  }, [rows, manpowerShiftsByKey, manpowerShiftsByDept, shiftOptions]);

  function shiftChoicesForRow(dept: string, jobSite: string, current: string) {
    const base = shiftChoiceBaseByRowKey.get(`${dept}||${jobSite}`) ?? shiftOptions;
    // dedupe by normalized shift key (กะ1 vs กะ 1) — prefer the row's own stored label for its key
    const byKey = new Map<string, string>();
    for (const label of base) byKey.set(normalizeShiftKey(label), label);
    if (current) byKey.set(normalizeShiftKey(current), current);
    return Array.from(byKey.values());
  }

  const deptOptions = Array.from(new Set(rows.map((r) => r.dept).filter(Boolean))).sort();

  // cascading: หน้างาน dropdown แสดงแค่ค่าที่มีในหน่วยงานที่เลือก
  const availableJobSiteOptions = (() => {
    const base = selectedDept === "all" ? rows : rows.filter((r) => r.dept === selectedDept);
    return Array.from(new Set(base.map((r) => r.jobSite).filter(Boolean))).sort();
  })();

  // cascading: หน่วยงาน dropdown แสดงแค่ค่าที่มีในหน้างานที่เลือก
  const availableDeptOptions = (() => {
    if (selectedJobSite === "all") return deptOptions;
    if (selectedJobSite === "__empty__") return Array.from(new Set(rows.filter((r) => r.jobSite === "").map((r) => r.dept).filter(Boolean))).sort();
    return Array.from(new Set(rows.filter((r) => r.jobSite === selectedJobSite).map((r) => r.dept).filter(Boolean))).sort();
  })();

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
    if (selectedJobSite !== "all") {
      if (selectedJobSite === "__empty__" && row.jobSite !== "") return false;
      if (selectedJobSite !== "__empty__" && row.jobSite !== selectedJobSite) return false;
    }
    if (!normalizedQuery) return true;
    return [row.empId, row.name, row.dept, row.jobSite, row.dayoff, row.shift]
      .some((v) => v.toLowerCase().includes(normalizedQuery));
  });

  const modifiedIds = new Set(
    rows
      .filter((row) => {
        const orig = originalRows.find((r) => r.id === row.id);
        return orig && (orig.dayoff !== row.dayoff || orig.shift !== row.shift || orig.shiftStart !== row.shiftStart || orig.jobSite !== row.jobSite);
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
        const allJobSites = new Set<string>();
        const skillCFMap = new Map<string, string>();
        for (const row of empRows) {
          const empId = cleanEmpId(
            row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"],
          );
          const dept = String(
            row["หน่วยงาน"] ?? row["Name (Section)"] ?? row["Department"] ?? "",
          ).trim();
          if (empId && dept) deptMap.set(empId, dept);
          const jobSite = String(row["หน่วยงานย่อย/Skill"] ?? row["หน้างาน"] ?? "").trim();
          if (empId && jobSite) {
            skillCFMap.set(empId, jobSite);
            allJobSites.add(jobSite);
          }
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
            const raw = r[colIdx];
            // ข้าม integer เพราะอาจเป็น skill level (0-5) ไม่ใช่ Excel time serial
            if (typeof raw === "number" && raw % 1 === 0) continue;
            const t = normalizeTimeText(raw);
            if (t) timeMap.set(empId, t);
          }
        }
        const parsed = dayoffRows.map((row, i) => {
          const r = toDayoffShiftEditorRow(row, i);
          if (!r.dept && deptMap.has(r.empId)) r.dept = deptMap.get(r.empId)!;
          if (!r.jobSite) r.jobSite = skillCFMap.get(r.empId) ?? "";
          if (r.jobSite) allJobSites.add(r.jobSite);
          if (!r.shiftStart && timeMap.has(r.empId)) r.shiftStart = timeMap.get(r.empId)!;
          if (!r.shiftEnd && r.shiftStart) r.shiftEnd = addHoursToTime(r.shiftStart, 9);
          return r;
        });
        setJobSiteOptions(Array.from(allJobSites).filter(Boolean).sort());
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

  function updateRow(id: string, field: "dayoff" | "shift" | "jobSite", value: string) {
    const targets =
      field === "dayoff"
        ? ["วันหยุดประจำสัปดาห์", "วันหยุด", "dayoff", "Dayoff", "Day Off"]
        : field === "shift"
        ? ["อยู่กะไหน", "shift", "กะ", "Shift"]
        : ["หน่วยงานย่อย/Skill", "หน้างาน", "job_site", "Job Site"];
    setRows((current) =>
      current.map((row) => {
        if (row.id !== id) return row;
        let raw = setRowCol(row.raw, value, ...targets);
        const next: typeof row = { ...row, [field]: value, raw };
        if (field === "shift" || field === "jobSite") {
          const newJobSite = field === "jobSite" ? value : row.jobSite;
          const newShift = field === "shift" ? value : row.shift;
          const locked = lookupManpower(row.dept, newJobSite, newShift);
          // ถ้าไม่พบเวลาใน Manpower สำหรับกะ/หน่วยงานย่อยใหม่ ต้องล้างเวลาเก่าทิ้ง ไม่ใช่ปล่อยให้ค้างเวลาของกะเดิมไว้คู่กับกะใหม่
          next.shiftStart = locked?.shiftStart ?? "";
          next.shiftEnd = locked?.shiftEnd ?? "";
          raw = setRowCol(raw, next.shiftStart, "เวลาเข้างาน", "เวลาเข้า", "shift_start");
          raw = setRowCol(raw, next.shiftEnd, "เวลาออก", "เวลาออกงาน", "shift_end");
          next.raw = raw;
        }
        return next;
      }),
    );
  }

  // แก้เวลาเข้า/ออกของพนักงาน 1 คนตรงในตาราง — เช็คกับ Manpower ก่อน ถ้ายังไม่มีเวลานี้บันทึกไว้
  // จะสร้างแถวใหม่ให้อัตโนมัติ (ตั้งชื่อหน่วยงานย่อยใหม่กันชนถ้าจำเป็น) แล้วผูกพนักงานคนนี้เข้ากับแถวนั้น
  function updateRowTime(row: DayoffShiftEditorRow, field: "shiftStart" | "shiftEnd", value: string) {
    if (!row.dept || !row.shift || row.shift === "ผู้จัดการ") return;
    const newStart = field === "shiftStart" ? value : row.shiftStart;
    const newEnd = field === "shiftEnd" ? value : row.shiftEnd;

    let finalJobSite = row.jobSite;
    if (newStart && newEnd) {
      const resolved = resolveTimeChange(row.dept, row.jobSite, row.shift, newStart, newEnd);
      finalJobSite = resolved.jobSite;
      if (resolved.isNew) addManpowerRowIfNeeded(row.dept, finalJobSite, row.shift, newStart, newEnd);
    }

    setRows((current) =>
      current.map((r) => {
        if (r.id !== row.id) return r;
        let raw = setRowCol(r.raw, finalJobSite, "หน่วยงานย่อย/Skill", "หน้างาน", "job_site", "Job Site");
        raw = setRowCol(raw, newStart, "เวลาเข้างาน", "เวลาเข้า", "shift_start");
        raw = setRowCol(raw, newEnd, "เวลาออก", "เวลาออกงาน", "shift_end");
        return { ...r, jobSite: finalJobSite, shiftStart: newStart, shiftEnd: newEnd, raw };
      }),
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
    const hasBulkTime = !!(bulkShiftStart && bulkShiftEnd);
    if (!bulkDayoff && !bulkShift && !bulkJobSite && !hasBulkTime) return;

    // ก่อน map ต้องเช็ค/สร้างแถว Manpower ให้ครบทุก dept+หน่วยงานย่อย+กะ ที่ไม่ซ้ำกันในกลุ่มที่เลือก (ทำครั้งเดียวต่อ 1 คู่ที่ไม่ซ้ำ)
    const resolvedJobSiteByKey = new Map<string, string>();
    if (hasBulkTime) {
      for (const row of rows) {
        if (!selectedIds.has(row.id)) continue;
        const jobSite = bulkJobSite || row.jobSite;
        const shift = bulkShift || row.shift;
        if (!row.dept || !shift || shift === "ผู้จัดการ") continue;
        const key = `${row.dept}||${jobSite}||${normalizeShiftKey(shift)}`;
        if (resolvedJobSiteByKey.has(key)) continue;
        const resolved = resolveTimeChange(row.dept, jobSite, shift, bulkShiftStart, bulkShiftEnd);
        resolvedJobSiteByKey.set(key, resolved.jobSite);
        if (resolved.isNew) addManpowerRowIfNeeded(row.dept, resolved.jobSite, shift, bulkShiftStart, bulkShiftEnd);
      }
    }

    setRows((current) =>
      current.map((row) => {
        if (!selectedIds.has(row.id)) return row;
        let raw = row.raw;
        let shiftStart = row.shiftStart;
        let shiftEnd = row.shiftEnd;
        let jobSite = bulkJobSite || row.jobSite;
        const shift = bulkShift || row.shift;

        if (bulkDayoff) raw = setRowCol(raw, bulkDayoff, "วันหยุดประจำสัปดาห์", "วันหยุด", "dayoff", "Dayoff", "Day Off");
        if (bulkShift) raw = setRowCol(raw, bulkShift, "อยู่กะไหน", "shift", "กะ", "Shift");

        if (hasBulkTime && row.dept && shift && shift !== "ผู้จัดการ") {
          const key = `${row.dept}||${jobSite}||${normalizeShiftKey(shift)}`;
          jobSite = resolvedJobSiteByKey.get(key) ?? jobSite;
          shiftStart = bulkShiftStart;
          shiftEnd = bulkShiftEnd;
        } else if (bulkShift || bulkJobSite) {
          const locked = lookupManpower(row.dept, jobSite, shift);
          // ถ้าไม่พบเวลาใน Manpower สำหรับหน่วยงานย่อย/กะใหม่ ต้องล้างเวลาเก่าทิ้ง ไม่ใช่ปล่อยให้ค้างเวลาของกะเดิมไว้คู่กับกะใหม่
          shiftStart = locked?.shiftStart ?? "";
          shiftEnd = locked?.shiftEnd ?? "";
        }

        if (jobSite !== row.jobSite) raw = setRowCol(raw, jobSite, "หน่วยงานย่อย/Skill", "หน้างาน", "job_site", "Job Site");
        if (shiftStart !== row.shiftStart) raw = setRowCol(raw, shiftStart, "เวลาเข้างาน", "เวลาเข้า", "shift_start");
        if (shiftEnd !== row.shiftEnd) raw = setRowCol(raw, shiftEnd, "เวลาออก", "เวลาออกงาน", "shift_end");

        return { ...row, dayoff: bulkDayoff || row.dayoff, shift, jobSite, shiftStart, shiftEnd, raw };
      }),
    );
    setSelectedIds(new Set());
    setBulkDayoff("");
    setBulkShift("");
    setBulkJobSite("");
    setBulkShiftStart("");
    setBulkShiftEnd("");
  }

  async function handleSave() {
    // นับคนที่มีกะแล้วแต่ยังไม่มีข้อมูลใน Manpower Plan (ไม่นับคนที่ยังไม่ได้ตั้งกะ หรือผู้จัดการซึ่งไม่มีกะตายตัวอยู่แล้ว)
    const unresolved = rows.filter(
      (row) => row.shift && row.shift !== "ผู้จัดการ" && !lookupManpower(row.dept, row.jobSite, row.shift),
    );
    if (unresolved.length > 0) {
      const ok = window.confirm(
        `พบ ${unresolved.length} คนที่หน่วยงาน + หน่วยงานย่อย + กะ ยังไม่มีข้อมูลใน Manpower Plan\n` +
        `เวลาเข้า-ออกของคนเหล่านี้จะถูกบันทึกเป็นค่าว่าง (แสดง — ในตาราง)\n\n` +
        `แนะนำ: ไปตั้งค่าที่หน้า Manpower Plan ก่อน แล้วค่อยกลับมาบันทึกหน้านี้\n\n` +
        `ต้องการบันทึกต่อเลยหรือไม่?`,
      );
      if (!ok) return;
    }
    setIsSaving(true);
    try {
      // บันทึก Manpower ก่อน ถ้ามีกะใหม่ที่สร้างจากการปรับเวลาในหน้านี้ค้างอยู่ — ต้องให้ Manpower มีแถวนั้นจริง
      // ก่อนที่ resolve เวลาด้านล่างจะไปค้นหามันเจอ
      if (manpowerDirty) {
        const mpRowsToSave: ManpowerEditorRow[] = manpowerRows.map((r, i) => ({ id: `mp-${i}`, ...r }));
        await saveManpowerRows(mpRowsToSave);
        setManpowerDirty(false);
      }
      const resolved = rows.map((row) => {
        // ผู้จัดการไม่มีกะตายตัว ไม่ต้อง resolve เวลา ปล่อยตามเดิม (มักจะว่างอยู่แล้ว)
        if (row.shift === "ผู้จัดการ") return row;
        const locked = lookupManpower(row.dept, row.jobSite, row.shift);
        // ถ้าไม่พบใน Manpower ต้องล้างเวลาทิ้ง ไม่ใช่ปล่อยเวลาเก่าที่อาจไม่ตรงกับกะปัจจุบันค้างไว้
        // (สอดคล้องกับข้อความแจ้งเตือนก่อนบันทึกด้านบน และกับ updateRow/applyBulk/ปุ่มสลับกะ)
        const shiftStart = locked?.shiftStart ?? "";
        const shiftEnd = locked?.shiftEnd ?? "";
        if (shiftStart === row.shiftStart && shiftEnd === row.shiftEnd) return row;
        let raw = setRowCol(row.raw, shiftStart, "เวลาเข้างาน", "เวลาเข้า", "shift_start");
        raw = setRowCol(raw, shiftEnd, "เวลาออก", "เวลาออกงาน", "shift_end");
        return { ...row, shiftStart, shiftEnd, raw };
      });
      await saveDayoffShiftRows(resolved);
      setRows(resolved);
      setOriginalRows(resolved);
    } catch {
      // error already set by saveDayoffShiftRows/saveManpowerRows
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
        <div style={{ display: "flex", gap: 8 }}>
          <button
            className="secondary-button"
            disabled={!rows.length}
            onClick={syncJobSitesFromManpower}
            type="button"
            title="เทียบหน่วยงานย่อยของแต่ละคนกับ Manpower Plan แล้วแนะนำค่าที่ควรจะเป็น"
          >
            <BarChart3 size={15} />
            แนะนำหน่วยงานย่อยจาก Manpower
          </button>
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
      </div>

      <div className="dayoff-editor-filters">
        <input
          aria-label="ค้นหา dayoff shift"
          placeholder="ค้นหา รหัส ชื่อ หน่วยงาน"
          type="search"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
        />
        <select value={selectedDept} onChange={(e) => {
          const val = e.target.value;
          setSelectedDept(val);
          // reset jobSite ถ้าค่าที่เลือกไม่มีในหน่วยงานใหม่
          if (selectedJobSite !== "all" && selectedJobSite !== "__empty__") {
            const nextJobSites = new Set(
              (val === "all" ? rows : rows.filter((r) => r.dept === val)).map((r) => r.jobSite)
            );
            if (!nextJobSites.has(selectedJobSite)) setSelectedJobSite("all");
          }
        }}>
          <option value="all">ทุกหน่วยงาน</option>
          {availableDeptOptions.map((d) => <option key={d} value={d}>{d}</option>)}
        </select>
        <select value={selectedJobSite} onChange={(e) => {
          const val = e.target.value;
          setSelectedJobSite(val);
          // reset dept ถ้าค่าที่เลือกไม่มีในหน้างานใหม่
          if (selectedDept !== "all") {
            const nextDepts = new Set(
              val === "all" ? rows.map((r) => r.dept) :
              val === "__empty__" ? rows.filter((r) => r.jobSite === "").map((r) => r.dept) :
              rows.filter((r) => r.jobSite === val).map((r) => r.dept)
            );
            if (!nextDepts.has(selectedDept)) setSelectedDept("all");
          }
        }}>
          <option value="all">ทุกหน่วยงานย่อย</option>
          <option value="__empty__">— ยังไม่มีข้อมูล</option>
          {availableJobSiteOptions.map((o) => <option key={o} value={o}>{o}</option>)}
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
        <button
          className="ghost-button"
          onClick={() => {
            setQuery("");
            setSelectedDept("all");
            setSelectedJobSite("all");
            setSelectedDayoff("all");
            setSelectedShift("all");
          }}
          type="button"
        >
          Clear
        </button>
      </div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 6 }}>
        <span className="dayoff-count">
          {filteredRows.length.toLocaleString()} / {rows.length.toLocaleString()} คน
          {modifiedIds.size > 0 && <span className="modified-badge">{modifiedIds.size} แก้ไข</span>}
        </span>
        {selectedDept !== "all" && selectedJobSite !== "all" && selectedJobSite !== "__empty__" && (
          <button
            className="primary-button"
            style={{ height: 30, fontSize: 12, padding: "0 14px", background: "linear-gradient(135deg, #f59e0b, #d97706)", gap: 6 }}
            type="button"
            onClick={() => {
              const กะ1Row = filteredRows.find((r) => r.shift === "กะ1");
              const กะ2Row = filteredRows.find((r) => r.shift === "กะ2");
              if (!กะ1Row || !กะ2Row) {
                alert("ต้องมีพนักงานทั้งกะ1 และ กะ2 ในมุมมองนี้จึงจะสลับได้");
                return;
              }
              const idsToSwap = new Set(filteredRows.filter((r) => r.shift === "กะ1" || r.shift === "กะ2").map((r) => r.id));
              setRows((current) =>
                current.map((row) => {
                  if (!idsToSwap.has(row.id)) return row;
                  const newShift = row.shift === "กะ1" ? "กะ2" : "กะ1";
                  const locked = lookupManpower(row.dept, row.jobSite, newShift);
                  let raw = setRowCol(row.raw, newShift, "อยู่กะไหน", "shift", "กะ", "Shift");
                  // ถ้าไม่พบเวลาใน Manpower สำหรับกะใหม่ ต้องล้างเวลาเก่าทิ้ง ไม่ใช่ปล่อยให้ค้างเวลาของกะเดิมไว้คู่กับกะใหม่
                  const shiftStart = locked?.shiftStart ?? "";
                  const shiftEnd = locked?.shiftEnd ?? "";
                  raw = setRowCol(raw, shiftStart, "เวลาเข้างาน", "เวลาเข้า", "shift_start");
                  raw = setRowCol(raw, shiftEnd, "เวลาออก", "เวลาออกงาน", "shift_end");
                  return { ...row, shift: newShift, shiftStart, shiftEnd, raw };
                })
              );
            }}
          >
            ⇄ สลับกะ 1 ↔ 2
          </button>
        )}
      </div>

      {selectedIds.size > 0 && (() => {
        const selectedDepts = Array.from(new Set(rows.filter((r) => selectedIds.has(r.id)).map((r) => r.dept)));
        const bulkJobSiteOptions = Array.from(
          new Set(selectedDepts.flatMap((d) => jobSiteChoicesForDept(d, ""))),
        ).sort();
        return (
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
          <select value={bulkJobSite} onChange={(e) => setBulkJobSite(e.target.value)}>
            <option value="">เปลี่ยนหน่วยงานย่อย...</option>
            {bulkJobSiteOptions.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
          <select value={bulkShift} onChange={(e) => setBulkShift(e.target.value)}>
            <option value="">เปลี่ยน Shift...</option>
            {shiftOptions.map((o) => <option key={o} value={o}>{o}</option>)}
          </select>
          <TimeInput24
            value={bulkShiftStart}
            onChange={setBulkShiftStart}
            title="เปลี่ยนเวลาเข้า"
            className="time24-btn shift-time-input"
          />
          <TimeInput24
            value={bulkShiftEnd}
            onChange={setBulkShiftEnd}
            title="เปลี่ยนเวลาออก"
            className="time24-btn shift-time-input"
          />
          <button
            className="primary-button"
            disabled={!bulkDayoff && !bulkShift && !bulkJobSite && !(bulkShiftStart && bulkShiftEnd)}
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
        );
      })()}

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
              <th>หน่วยงาน</th>
              <th>หน่วยงานย่อย/Skill</th>
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
                  <select value={row.jobSite} onChange={(e) => updateRow(row.id, "jobSite", e.target.value)}>
                    <option value="">-</option>
                    {jobSiteChoicesForDept(row.dept, row.jobSite).map((o) => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>
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
                    {shiftChoicesForRow(row.dept, row.jobSite, row.shift).map((o) => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>
                {(() => {
                  if (row.shift === "ผู้จัดการ") {
                    return (
                      <>
                        <td><span style={{ color: "#94a3b8", fontSize: 13 }}>—</span></td>
                        <td className="shift-end-cell"><span style={{ color: "#94a3b8", fontSize: 13 }}>—</span></td>
                      </>
                    );
                  }
                  const locked = lookupManpower(row.dept, row.jobSite, row.shift);
                  const knownTitle = locked
                    ? "ตรงกับเวลาที่ตั้งไว้ใน Manpower"
                    : "ยังไม่มีเวลานี้ใน Manpower — กด Save แล้วระบบจะบันทึกเป็นกะใหม่ให้อัตโนมัติ";
                  return (
                    <>
                      <td>
                        <TimeInput24
                          value={row.shiftStart}
                          onChange={(v) => updateRowTime(row, "shiftStart", v)}
                          className={`time24-btn shift-time-input${locked ? "" : " time24-unresolved"}`}
                          title={knownTitle}
                        />
                      </td>
                      <td className="shift-end-cell">
                        <TimeInput24
                          value={row.shiftEnd}
                          onChange={(v) => updateRowTime(row, "shiftEnd", v)}
                          className={`time24-btn shift-time-input${locked ? "" : " time24-unresolved"}`}
                          title={knownTitle}
                        />
                      </td>
                    </>
                  );
                })()}
              </tr>
            ))}
            {!activeFile ? (
              <tr><td colSpan={9}>อัปโหลด Dayoff & Shift master ก่อน จึงจะแก้ไขในหน้านี้ได้</td></tr>
            ) : null}
            {activeFile && isLoading ? (
              <tr><td colSpan={9}>Loading Dayoff & Shift...</td></tr>
            ) : null}
            {activeFile && !isLoading && filteredRows.length === 0 ? (
              <tr><td colSpan={9}>ไม่พบข้อมูลที่ค้นหา</td></tr>
            ) : null}
          </tbody>
        </table>
      </div>
      {filteredRows.some((r) => {
        const locked = lookupManpower(r.dept, r.jobSite, r.shift);
        const start = locked?.shiftStart || r.shiftStart;
        const end = locked?.shiftEnd || r.shiftEnd || (start ? addHoursToTime(start, 9) : "");
        return end && Number(end.split(":")[0]) < Number((start || "0").split(":")[0]);
      }) && (
        <p style={{ fontSize: 12, color: "#94a3b8", marginTop: 6 }}>
          <span style={{ color: "#fcd34d", fontWeight: 600 }}>+1D</span> = เวลาออกงานเป็นวันถัดไป (กะกลางคืน)
        </p>
      )}
    </section>
  );
}

function SkillMatrixPage({
  activeFile,
  dayoffShiftFile,
  employeeMasterFile,
  saveSkillMatrixRows,
}: {
  activeFile?: MasterFile;
  dayoffShiftFile?: MasterFile;
  employeeMasterFile?: MasterFile;
  saveSkillMatrixRows: (rows: SkillMatrixSaveRow[]) => Promise<void>;
}) {
  const [rows, setRows] = useState<SkillFlatRow[]>([]);
  const [empInfoMap, setEmpInfoMap] = useState<Map<string, { name: string; dept: string; jobSite: string }>>(new Map());
  const [allEmpList, setAllEmpList] = useState<Array<{ empId: string; name: string }>>([]);
  const [query, setQuery] = useState("");
  const [selectedDept, setSelectedDept] = useState("all");
  const [selectedSkill, setSelectedSkill] = useState("all");
  const [selectedShift, setSelectedShift] = useState("all");
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [bulkLevel, setBulkLevel] = useState("");
  const [addEmpId, setAddEmpId] = useState("");
  const [addSkill, setAddSkill] = useState("");
  const [addLevel, setAddLevel] = useState(1);
  const [addOpen, setAddOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [isSaving, setIsSaving] = useState(false);

  const LEVEL_LABELS: Record<number, string> = { 1: "น้อย", 2: "ปานกลาง", 3: "ถนัด" };
  const LEVEL_BG: Record<number, string> = { 1: "#fecaca", 2: "#fef08a", 3: "#34d399", 4: "#93c5fd", 5: "#a78bfa" };

  useEffect(() => {
    if (!employeeMasterFile?.file_path && !activeFile?.file_path) {
      setRows([]);
      setEmpInfoMap(new Map());
      setAllEmpList([]);
      setSelectedIds(new Set());
      return;
    }

    let isMounted = true;
    setIsLoading(true);

    const skillPromise = activeFile?.file_path
      ? downloadSheetRows(activeFile.file_path).catch(() => [] as Record<string, unknown>[])
      : Promise.resolve([] as Record<string, unknown>[]);
    const empPromise = employeeMasterFile?.file_path
      ? downloadSheetRows(employeeMasterFile.file_path).catch(() => [] as Record<string, unknown>[])
      : Promise.resolve([] as Record<string, unknown>[]);
    const dayoffPromise = dayoffShiftFile?.file_path
      ? downloadSheetRows(dayoffShiftFile.file_path).catch(() => [] as Record<string, unknown>[])
      : Promise.resolve([] as Record<string, unknown>[]);

    Promise.all([skillPromise, empPromise, dayoffPromise])
      .then(([skillRows, empRows, dayoffRows]) => {
        if (!isMounted) return;

        const empInfo = new Map<string, { name: string; dept: string; jobSite: string }>();
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
          const jobSite = String(row["หน่วยงานย่อย/Skill"] ?? row["หน้างาน"] ?? "").trim();
          if (empId) empInfo.set(empId, { name, dept, jobSite });
        }

        const dayoffMap = buildDayoffShiftMap(dayoffRows);

        const parsed: SkillFlatRow[] = skillRows
          .map((row, i) => {
            const empId = cleanEmpId(
              row["Employee ID"] ?? row["Emp ID"] ?? row["emp_id"] ?? row["รหัสพนักงาน"],
            );
            const skill = String(row["Skill"] ?? row["skill"] ?? row["ทักษะ"] ?? "").trim();
            const level = Number(row["Level"] ?? row["level"] ?? row["ระดับ"]) || 0;
            const info = empInfo.get(empId) ?? { name: empId, dept: "", jobSite: "" };
            const ds = dayoffMap.get(empId) ?? { dayoff: "", shift: "", shiftStart: "" };
            return {
              id: `${i}-${empId}-${skill}`,
              empId,
              name: info.name,
              dept: info.dept,
              jobSite: info.jobSite,
              shift: ds.shift,
              shiftStart: ds.shiftStart,
              dayoff: ds.dayoff,
              skill,
              level,
              origLevel: level,
            };
          })
          .filter((r) => r.empId && r.skill);

        const empList = Array.from(empInfo.entries()).map(([empId, info]) => ({ empId, name: info.name }));
        setRows(parsed);
        setEmpInfoMap(empInfo);
        setAllEmpList(empList);
        setSelectedIds(new Set());
      })
      .catch(() => {
        if (!isMounted) return;
        setRows([]);
        setEmpInfoMap(new Map());
        setAllEmpList([]);
      })
      .finally(() => {
        if (!isMounted) return;
        setIsLoading(false);
      });

    return () => { isMounted = false; };
  }, [activeFile?.file_path, employeeMasterFile?.file_path, dayoffShiftFile?.file_path]);

  const deptOptions = Array.from(new Set(rows.map((r) => r.dept).filter(Boolean))).sort();
  const skillOptions = Array.from(new Set(rows.map((r) => r.skill).filter(Boolean))).sort();
  const shiftOptions = Array.from(new Set(rows.map((r) => r.shift).filter(Boolean))).sort();
  const normalizedQuery = query.trim().toLowerCase();
  const filteredRows = rows.filter((row) => {
    if (selectedDept !== "all" && row.dept !== selectedDept) return false;
    if (selectedSkill !== "all" && row.skill !== selectedSkill) return false;
    if (selectedShift !== "all" && row.shift !== selectedShift) return false;
    if (!normalizedQuery) return true;
    return [row.empId, row.name, row.dept, row.skill, row.shift, row.dayoff, row.jobSite].some((v) =>
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
    } catch {
      // error already set by saveSkillMatrixRows
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
      const empData = empInfoMap.get(empId);
      const existingRow = rows.find((r) => r.empId === empId);
      setRows((prev) => [
        ...prev,
        {
          id: `add-${Date.now()}-${empId}-${skill}`,
          empId,
          name: empData?.name ?? existingRow?.name ?? empId,
          dept: empData?.dept ?? existingRow?.dept ?? "",
          jobSite: empData?.jobSite ?? existingRow?.jobSite ?? "",
          shift: existingRow?.shift ?? "",
          shiftStart: existingRow?.shiftStart ?? "",
          dayoff: existingRow?.dayoff ?? "",
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

  const levelBg = LEVEL_BG;

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

      <div className="table-filters dayoff-editor-filters" style={{ gridTemplateColumns: "1fr 160px 160px 140px auto" }}>
        <input
          aria-label="ค้นหา skill matrix"
          placeholder="ค้นหา รหัส ชื่อ แผนก กะ วันหยุด ทักษะ"
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
        <select value={selectedShift} onChange={(e) => setSelectedShift(e.target.value)}>
          <option value="all">ทุกกะ</option>
          {shiftOptions.map((s) => <option key={s} value={s}>{s}</option>)}
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
            <option value="0">— ยังไม่ระบุ</option>
            <option value="1">1 — น้อย</option>
            <option value="2">2 — ปานกลาง</option>
            <option value="3">3 — ถนัด</option>
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
              <th>หน่วยงานย่อย</th>
              <th>กะ</th>
              <th>เวลาเข้า</th>
              <th>วันหยุด</th>
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
                <td>{row.jobSite || "—"}</td>
                <td>{row.shift || "—"}</td>
                <td>{row.shiftStart || "—"}</td>
                <td>{row.dayoff || "—"}</td>
                <td>{row.skill}</td>
                <td>
                  <select
                    className="skill-level-cell"
                    style={{ background: levelBg[row.level] ?? "" }}
                    value={row.level}
                    onChange={(e) => updateRow(row.id, Number(e.target.value))}
                  >
                    <option value={0}>— ยังไม่ระบุ</option>
                    <option value={1}>1 — น้อย</option>
                    <option value={2}>2 — ปานกลาง</option>
                    <option value={3}>3 — ถนัด</option>
                    {row.level > 3 && <option value={row.level}>{row.level} — (เก่า)</option>}
                  </select>
                </td>
              </tr>
            ))}
            {!employeeMasterFile && !activeFile ? (
              <tr><td colSpan={10} style={{ color: "var(--muted)", textAlign: "center", padding: "24px 0" }}>อัปโหลดไฟล์รายชื่อพนักงาน (Master Data → Files) เพื่อเริ่มต้นใช้งาน</td></tr>
            ) : null}
            {isLoading ? (
              <tr><td colSpan={10} style={{ textAlign: "center", padding: "24px 0" }}>กำลังโหลด...</td></tr>
            ) : null}
            {!isLoading && (employeeMasterFile || activeFile) && filteredRows.length === 0 ? (
              <tr><td colSpan={10} style={{ color: "var(--muted)", textAlign: "center", padding: "24px 0" }}>ยังไม่มี Skill — กด "+ เพิ่ม Skill ให้พนักงาน" เพื่อเริ่มเพิ่มข้อมูล</td></tr>
            ) : null}
          </tbody>
        </table>
      </div>

      {/* Add row */}
      {(activeFile || employeeMasterFile) && (
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
                {(allEmpList.length > 0 ? allEmpList : Array.from(new Set(rows.map((r) => r.empId))).map((id) => ({ empId: id, name: rows.find((x) => x.empId === id)?.name ?? id }))).map(({ empId, name }) => (
                  <option key={empId} value={empId}>{name}</option>
                ))}
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
                {([1, 2, 3] as const).map((v) => (
                  <option key={v} value={v}>{v} — {LEVEL_LABELS[v]}</option>
                ))}
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

const STORAGE_WARN_MB = 1001;
const STORAGE_LIMIT_MB = 1024;

function estimateStorageMB(runs: AllocationRun[]): number {
  const bytes = runs.reduce((acc, r) => acc + (r.record_count ?? 500) * 200, 0);
  return Math.round(bytes / (1024 * 1024));
}

function TimestampPage({
  createDailyRun,
  deleteRun,
  downloadTimestampFile,
  hasAllActiveMasters,
  isCreatingRun,
  isLoadingReport,
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
  isLoadingReport: boolean;
  latestRun?: AllocationRun;
  runs: AllocationRun[];
  setTimestampFile: (file: File | null) => void;
  timestampFile: File | null;
}) {
  const [pendingDelete, setPendingDelete] = useState<AllocationRun | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  const [showStorageWarn, setShowStorageWarn] = useState(false);

  const estimatedMB = estimateStorageMB(runs);
  const storageUsedPct = Math.min(100, Math.round((estimatedMB / STORAGE_LIMIT_MB) * 100));
  const isStorageWarning = estimatedMB >= STORAGE_WARN_MB;
  const oldestRun = runs.length > 0
    ? [...runs].sort((a, b) => a.created_at.localeCompare(b.created_at))[0]
    : null;

  async function handleUploadClick() {
    if (isStorageWarning) { setShowStorageWarn(true); return; }
    await createDailyRun();
  }

  async function confirmDelete() {
    if (!pendingDelete) return;
    setIsDeleting(true);
    await deleteRun(pendingDelete);
    setIsDeleting(false);
    setPendingDelete(null);
  }

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

        {isStorageWarning && (
          <div className="storage-warn-banner">
            <AlertTriangle size={15} />
            <span>พื้นที่เริ่มเต็ม · {runs.length} ไฟล์ · แนะนำให้ลบไฟล์เก่าก่อน upload</span>
          </div>
        )}

        <button
          className="primary-button ts-submit-btn"
          disabled={!timestampFile || !hasAllActiveMasters || isCreatingRun}
          onClick={() => void handleUploadClick()}
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

        {runs.length > 0 && (
          <div className="storage-bar-row">
            <div className="storage-bar-track">
              <div
                className={`storage-bar-fill${storageUsedPct >= 80 ? " danger" : storageUsedPct >= 60 ? " warn" : ""}`}
                style={{ width: `${storageUsedPct}%` }}
              />
            </div>
            <span className="storage-bar-label">
              ~{estimatedMB} MB / {STORAGE_LIMIT_MB} MB ({storageUsedPct}%)
            </span>
          </div>
        )}

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
            const isOldest = oldestRun?.id === run.id;
            return (
              <div className={`ts-history-row${isOldest && isStorageWarning ? " ts-row-oldest" : ""}`} key={run.id}>
                <div className="ts-history-info">
                  <strong>{filename}</strong>
                  <span>{meta}{isOldest && isStorageWarning ? <span className="ts-oldest-badge">เก่าที่สุด</span> : null}</span>
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
                    onClick={() => setPendingDelete(run)}
                    title={isLoadingReport ? "กำลังโหลด report — รอสักครู่" : "ลบ"}
                    disabled={isLoadingReport}
                    type="button"
                  >
                    <Trash2 size={15} />
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      </section>

      {/* Storage warning modal */}
      {showStorageWarn && (
        <div className="modal-overlay" onClick={() => setShowStorageWarn(false)}>
          <div className="modal-box" onClick={(e) => e.stopPropagation()}>
            <div className="modal-icon warn"><AlertTriangle size={28} /></div>
            <h3>พื้นที่จัดเก็บใกล้เต็ม</h3>
            <p className="modal-desc">
              ขณะนี้มี <strong>{runs.length} ไฟล์</strong> ในระบบ (ประมาณ ~{estimatedMB} MB)
              แนะนำให้ลบไฟล์เก่าออกก่อน upload ใหม่
            </p>
            {oldestRun && (
              <div className="modal-suggest">
                <span className="modal-suggest-label">ไฟล์เก่าที่สุด</span>
                <strong>{oldestRun.original_filename ?? oldestRun.id}</strong>
                <span>{new Date(oldestRun.created_at).toLocaleDateString("th-TH", { day: "numeric", month: "long", year: "numeric" })}</span>
                <button
                  className="danger-button"
                  type="button"
                  onClick={() => { setShowStorageWarn(false); setPendingDelete(oldestRun); }}
                >
                  <Trash2 size={14} /> ลบไฟล์นี้
                </button>
              </div>
            )}
            <div className="modal-actions">
              <button className="secondary-button" type="button" onClick={() => setShowStorageWarn(false)}>
                ยกเลิก
              </button>
              <button
                className="primary-button"
                type="button"
                onClick={() => { setShowStorageWarn(false); void createDailyRun(); }}
              >
                Upload ต่อไปเลย
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete confirm dialog */}
      {pendingDelete && (
        <div className="modal-overlay" onClick={() => !isDeleting && setPendingDelete(null)}>
          <div className="modal-box" onClick={(e) => e.stopPropagation()}>
            <div className="modal-icon danger"><Trash2 size={28} /></div>
            <h3>ยืนยันการลบไฟล์</h3>
            <p className="modal-desc">
              คุณต้องการลบไฟล์นี้ใช่ไหม? การลบไม่สามารถย้อนกลับได้
            </p>
            <div className="modal-file-info">
              <strong>{pendingDelete.original_filename ?? pendingDelete.id}</strong>
              <span>
                {pendingDelete.record_count != null ? `${pendingDelete.record_count.toLocaleString()} รายการ · ` : ""}
                {new Date(pendingDelete.created_at).toLocaleDateString("th-TH", { day: "numeric", month: "long", year: "numeric" })}
              </span>
            </div>
            <div className="modal-actions">
              <button
                className="secondary-button"
                type="button"
                disabled={isDeleting}
                onClick={() => setPendingDelete(null)}
              >
                ยกเลิก
              </button>
              <button
                className="danger-button"
                type="button"
                disabled={isDeleting}
                onClick={() => void confirmDelete()}
              >
                <Trash2 size={14} />
                {isDeleting ? "กำลังลบ..." : "ยืนยันลบ"}
              </button>
            </div>
          </div>
        </div>
      )}
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

  useEffect(() => { setPage(1); }, [reportData]);

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

  function handleExport() {
    const isoDate = reportData?.isoTargetDate ?? "";
    const data = sortedRows.map((row) => ({
      "รหัส": row.empId,
      "ชื่อ-สกุล": row.name,
      "หน่วยงาน": row.dept,
      "ตำแหน่ง": row.position,
      "กะ": row.shift,
      "เริ่มกะ": row.shiftStart,
      "Scan In": row.scanInDate ? `${row.scanInDate} ${row.scanIn}` : row.scanIn,
      "Scan Out": row.scanOutDate ? `${row.scanOutDate} ${row.scanOut}` : row.scanOut,
      "สถานะ": STATUS_TH[row.status] ?? row.status,
      "สาย": row.status === "Late" ? row.minutesLate : 0,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Timestamp");
    XLSX.writeFile(wb, `timestamp_${isoDate || "export"}.xlsx`);
  }

  return (
    <section className="panel results-panel">
      <div className="ot-detail-hdr">
        <h3>Timestamp With Dept<span className="ot-detail-count"> ({allRows.length} rows)</span></h3>
        <div className="ot-detail-filters">
          <input
            className="ot-detail-search"
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
            <option value="NoScanIn">ขาดสแกนเข้า</option>
            <option value="Pending">รอเข้างาน</option>
          </select>
          <button
            className="ghost-button"
            onClick={() => updateFilter(() => { setQuery(""); setDeptFilter("all"); setStatusFilter("all"); })}
            type="button"
          >
            Clear
          </button>
          <button type="button" className="primary-button small" onClick={handleExport} disabled={allRows.length === 0}>
            <Download size={14} />
            Export Excel
          </button>
        </div>
      </div>
      <div className="table-scroll">
        <table className="table data-table ts-dept-table">
          <thead>
            <tr>
              <th><SortButton columnKey="empId" setSort={setSort} sort={sort}>รหัส</SortButton></th>
              <th><SortButton columnKey="name" setSort={setSort} sort={sort}>ชื่อ-สกุล</SortButton></th>
              <th><SortButton columnKey="dept" setSort={setSort} sort={sort}>หน่วยงาน</SortButton></th>
              <th><SortButton columnKey="shift" setSort={setSort} sort={sort}>กะ</SortButton></th>
              <th><SortButton columnKey="shiftStart" setSort={setSort} sort={sort}>เริ่มกะ</SortButton></th>
              <th><SortButton columnKey="scanIn" setSort={setSort} sort={sort}>Scan In</SortButton></th>
              <th><SortButton columnKey="scanOut" setSort={setSort} sort={sort}>Scan Out</SortButton></th>
              <th><SortButton columnKey="status" setSort={setSort} sort={sort}>สถานะ</SortButton></th>
              <th><SortButton columnKey="minutesLate" setSort={setSort} sort={sort}>สาย</SortButton></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={`${row.empId}-${row.scanIn}`}>
                <td className="ts-col-id">{row.empId}</td>
                <td className="ts-col-name" title={row.name}>{row.name}</td>
                <td className="ts-col-dept" title={row.dept}>{row.dept}</td>
                <td>{row.shift}</td>
                <td>{row.shiftStart}</td>
                <td className="scan-cell">{scanDateBadge(row.scanInDate)}{row.scanIn}</td>
                <td className="scan-cell">{scanDateBadge(row.scanOutDate)}{row.scanOut}</td>
                <td><span className={`status-pill ${row.status.toLowerCase()}`}>{STATUS_TH[row.status] ?? row.status}</span></td>
                <td className="ts-col-late">{row.status === "Late" ? formatLateTime(row.minutesLate) : "-"}</td>
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
      <div className="ot-detail-hdr">
        <h3>ผลลัพธ์การจัดสรรล่าสุด<span className="ot-detail-count"> ({allRows.length} คน)</span></h3>
        <div className="ot-detail-filters">
          {standalone && <>
            <input
              className="ot-detail-search"
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
              onClick={() => updateFilter(() => { setQuery?.(""); setDeptFilter?.("all"); setStatusFilter?.("all"); })}
              type="button"
            >
              Clear
            </button>
          </>}
          <button className="ghost-button" type="button">ดูทั้งหมด</button>
          <button className="primary-button small" type="button">
            <Download size={14} /> Export <ChevronDown size={13} />
          </button>
        </div>
      </div>
      <div className="table-scroll">
        <table className="table data-table results-table">
          <thead>
            <tr>
              <th>No.</th>
              <th><SortButton columnKey="empId" setSort={setSort} sort={sort}>รหัสพนักงาน</SortButton></th>
              <th><SortButton columnKey="name" setSort={setSort} sort={sort}>ชื่อ-สกุล</SortButton></th>
              <th><SortButton columnKey="dept" setSort={setSort} sort={sort}>หน่วยงาน</SortButton></th>
              <th><SortButton columnKey="position" setSort={setSort} sort={sort}>ตำแหน่ง</SortButton></th>
              <th>หน่วยงานย่อย</th>
              <th>กะ</th>
              <th><SortButton columnKey="scanIn" setSort={setSort} sort={sort}>เวลาเข้า</SortButton></th>
              <th><SortButton columnKey="status" setSort={setSort} sort={sort}>สถานะ</SortButton></th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, index) => (
              <tr key={`${row.empId}-${row.scanIn}`}>
                <td>{(safePage - 1) * pageSize + index + 1}</td>
                <td>{row.empId}</td>
                <td className="res-col-name" title={row.name}>{row.name}</td>
                <td className="res-col-dept" title={row.dept}>{row.dept}</td>
                <td className="res-col-pos" title={row.position}>{row.position}</td>
                <td className="res-col-section" title={row.section || "-"}>{row.section || "-"}</td>
                <td>{row.shift || "-"}</td>
                <td>{row.scanIn}</td>
                <td>
                  <span className={`status-pill ${row.status.toLowerCase()}`}>
                    {STATUS_TH[row.status] ?? row.status}
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
  const [leaveError, setLeaveError] = useState("");

  const isoDate = reportData?.isoTargetDate ?? "";
  const [deptConfirmations, setDeptConfirmations] = useState<Map<string, { confirmed_by: string; confirmed_at: string }>>(new Map());
  const [confirmPanelCollapsed, setConfirmPanelCollapsed] = useState(true);
  const [lateAbsentCollapsed, setLateAbsentCollapsed] = useState(true);

  useEffect(() => {
    if (!isoDate) return;
    supabase.from("leave_records").select("emp_id, leave_type").eq("leave_date", isoDate)
      .then(({ data: rows }) => {
        if (!rows) return;
        const map = new Map(rows.map((r: { emp_id: string; leave_type: string }) => [r.emp_id, r.leave_type]));
        const absentIds = (reportData?.records ?? [])
          .filter((r: { status: string }) => r.status === "Absent")
          .map((r: { empId: string }) => r.empId);
        absentIds.filter((id: string) => !map.has(id)).forEach((id: string) => map.set(id, "ขาดงาน"));
        setLeaveMap(map);
      });
    supabase.from("daily_confirmations")
      .select("dept, confirmed_by, confirmed_at")
      .eq("confirm_date", isoDate)
      .then(({ data: rows }) => {
        if (!rows) return;
        setDeptConfirmations(new Map(rows.map((r: { dept: string; confirmed_by: string; confirmed_at: string }) => [r.dept, { confirmed_by: r.confirmed_by, confirmed_at: r.confirmed_at }])));
      });
  }, [isoDate]); // eslint-disable-line react-hooks/exhaustive-deps

  const saveLeave = async (empId: string, leaveType: string) => {
    if (!isoDate || !leaveType) return;
    const prev = leaveMap.get(empId);
    setLeaveMap(m => new Map(m).set(empId, leaveType));
    const { error } = await supabase.from("leave_records").upsert(
      { emp_id: empId, leave_date: isoDate, leave_type: leaveType },
      { onConflict: "emp_id,leave_date" }
    );
    if (error) {
      setLeaveMap(m => { const n = new Map(m); prev === undefined ? n.delete(empId) : n.set(empId, prev); return n; });
      setLeaveError(`บันทึกการลาไม่สำเร็จ: ${error.message}`);
      setTimeout(() => setLeaveError(""), 4000);
    }
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
    unmatchedScanIds: [],
    isoTargetDate: "",
    targetMonthKey: "",
  };

  const scopedRecords = selectedDept === "all"
    ? data.records
    : data.records.filter((row) => row.dept === selectedDept);
  const scopedTotal = scopedRecords.length;
  const scopedPresent = scopedRecords.filter((row) => row.status === "Present" || row.status === "NoScanIn").length;
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
  const presentPercent = activeTotal ? (scopedPresent / activeTotal) * 100 : 0;
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
  const absentRecords = scopedRecords.filter((r) => r.status === "Absent");
  const absTotalAbsent = absentRecords.length;
  const absConfirmed = absentRecords.filter((r) => { const lt = leaveMap.get(r.empId); return lt && lt !== "ขาดงาน"; }).length;
  const absPending = absTotalAbsent - absConfirmed;
  const absPct = absTotalAbsent > 0 ? Math.round((absConfirmed / absTotalAbsent) * 100) : 0;
  const absDeptMap = new Map<string, { total: number; confirmed: number }>();
  for (const r of absentRecords) {
    const s = absDeptMap.get(r.dept) ?? { total: 0, confirmed: 0 };
    s.total++;
    const lt = leaveMap.get(r.empId);
    if (lt && lt !== "ขาดงาน") s.confirmed++;
    absDeptMap.set(r.dept, s);
  }
  const absDeptRows = Array.from(absDeptMap.entries())
    .map(([dept, s]) => ({ dept, ...s, pending: s.total - s.confirmed }))
    .sort((a, b) => b.pending - a.pending);
  const lateDeptRows = Array.from(
    data.lateRows.reduce((map, row) => {
      map.set(row.dept, (map.get(row.dept) ?? 0) + 1);
      return map;
    }, new Map<string, number>()),
  )
    .map(([dept, count]) => ({ dept, count }))
    .sort((a, b) => b.count - a.count);
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
      {leaveError && (
        <div className="error-bar" style={{ marginBottom: 12 }}>{leaveError}</div>
      )}
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
          label="ขาด/ลา"
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
            <span><i className="absent" />ขาด/ลา</span>
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

        {absTotalAbsent > 0 && (
        <div className="panel report-overview-card">
          <div className="sup-report-header">
            <h3>สถานะการขาด/ลา{selectedDept !== "all" ? ` · ${selectedDept}` : ""}</h3>
            <span className={`sup-overall-badge ${absPending === 0 ? "done" : "pending"}`}>
              {absPending === 0 ? "✓ ครบแล้ว" : `⚠ ค้าง ${absPending}`}
            </span>
          </div>
          <div className="sup-report-progress">
            <div className="sup-progress-bar">
              <div className="sup-progress-fill" style={{ width: `${absPct}%`, background: "#ef4444" }} />
            </div>
            <span className="sup-progress-pct">{absPct}%</span>
          </div>
          <div className="sup-report-stats">
            <div className="sup-report-stat-item">
              <span className="sup-report-stat-val" style={{ color: "#10b981" }}>{absConfirmed}</span>
              <span className="sup-report-stat-lbl">บันทึกแล้ว</span>
            </div>
            <div className="sup-report-stat-item">
              <span className="sup-report-stat-val" style={{ color: "#ef4444" }}>{absTotalAbsent}</span>
              <span className="sup-report-stat-lbl">ขาด/ลาทั้งหมด</span>
            </div>
            <div className="sup-report-stat-item">
              <span className="sup-report-stat-val" style={{ color: absPending > 0 ? "#ef4444" : "#94a3b8" }}>{absPending}</span>
              <span className="sup-report-stat-lbl">ยังค้าง</span>
            </div>
          </div>
          <div className="sup-report-dept-list">
            {absDeptRows.map((row) => {
              const dPct = row.total > 0 ? Math.round((row.confirmed / row.total) * 100) : 100;
              const done = row.pending === 0;
              return (
                <div key={row.dept} className={`sup-dept-row ${done ? "done" : "pending"}`}>
                  <span className="sup-dept-name">{row.dept}</span>
                  <div className="sup-dept-bar-wrap">
                    <div className="sup-mini-bar">
                      <div className="sup-mini-fill" style={{ width: `${dPct}%`, background: "#ef4444" }} />
                    </div>
                    <span className="sup-mini-pct">{dPct}%</span>
                  </div>
                  <span className="sup-dept-counts">{row.confirmed}/{row.total}</span>
                  <span className={`sup-status-badge ${done ? "done" : "pending"}`}>
                    {done ? "✓" : `${row.pending}`}
                  </span>
                </div>
              );
            })}
          </div>
        </div>
        )}
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
                          <option value="ลาตรวจครรภ์">ลาตรวจครรภ์</option>
                          <option value="ลาคลอด">ลาคลอด</option>
                          <option value="ลาอุบัติเหตุจากการปฏิบัติงาน">ลาอุบัติเหตุจากการปฏิบัติงาน</option>
                          <option value="ลาบวช/ลาพิธีสำคัญทางศาสนา">ลาบวช/ลาพิธีสำคัญทางศาสนา</option>
                          <option value="ลาทหาร">ลาทหาร</option>
                          <option value="ลาพิเศษไม่จ่าย">ลาพิเศษไม่จ่าย</option>
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
                          <td className="muted-text">{conf ? new Date(conf.confirmed_at).toLocaleString("th-TH", { day: "numeric", month: "numeric", year: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false }) : "—"}</td>
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

      {data.unmatchedScanIds.length > 0 ? (
        <section className="panel report-card">
          <div className="panel-title-row">
            <h3>พบรหัสที่ scan แล้วแต่ไม่มีใน master ({data.unmatchedScanIds.length} คน)</h3>
          </div>
          <p className="empty-copy">
            คนเหล่านี้ scan เข้าออกจริง แต่รหัสพนักงานไม่ตรงกับไฟล์ master ที่ใช้อยู่ —
            ตรวจสอบว่าเป็นรหัสใหม่ที่ยังไม่ sync เข้า master หรือรูปแบบรหัสไม่ตรงกัน
          </p>
          <table className="table data-table">
            <thead>
              <tr>
                <th>รหัสพนักงาน</th>
                <th>ชื่อ (จากไฟล์ scan)</th>
                <th>เวลา scan แรกสุด</th>
              </tr>
            </thead>
            <tbody>
              {data.unmatchedScanIds.map((row) => (
                <tr key={row.empId}>
                  <td>{row.empId}</td>
                  <td>{row.name}</td>
                  <td>{row.scanIn}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </section>
      ) : null}
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
  onClick,
}: {
  icon: ReactNode;
  tone: "green" | "blue" | "amber" | "purple" | "gray" | "red";
  label: string;
  value: string;
  unit: string;
  note: string;
  progress?: number;
  onClick?: () => void;
}) {
  return (
    <article
      className={`kpi-card kpi-${tone}${onClick ? " kpi-clickable" : ""}`}
      onClick={onClick}
      role={onClick ? "button" : undefined}
      tabIndex={onClick ? 0 : undefined}
      onKeyDown={onClick ? (e) => { if (e.key === "Enter" || e.key === " ") onClick(); } : undefined}
    >
      <div className={`kpi-icon ${tone}`}>{icon}</div>
      <div className="kpi-body">
        <span>{label}</span>
        <div>
          <strong>{value}</strong>
          <b>{unit}</b>
        </div>
        <p>{note}</p>
        {progress != null ? (
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

function scanDateBadge(date: string) {
  if (!date) return null;
  const parts = date.split("-");
  if (parts.length < 3) return null;
  return (
    <span className="scan-date-badge" title={date}>
      {Number(parts[2])}/{Number(parts[1])}
    </span>
  );
}

function formatDateTH(isoDate: string) {
  const parts = isoDate?.split("-");
  if (!parts || parts.length < 3) return isoDate ?? "-";
  const m = Number(parts[1]);
  const d = Number(parts[2]);
  if (isNaN(m) || isNaN(d) || m < 1 || m > 12) return isoDate;
  const thaiMonths = ["", "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  return `${d} ${thaiMonths[m]}`;
}

// ─────────────────────────────────────────────────────────────────────────────
// OT Dashboard
// ─────────────────────────────────────────────────────────────────────────────

// หน่วยงานที่มีเวลาพัก 1 ชั่วโมงหลังเลิกงาน (หน่วยงานอื่นพัก 30 นาที) — ช่วงพักนี้ไม่นับเป็น OT
const OT_LONG_BREAK_DEPTS = new Set(["งานตัดแต่งพิเศษ", "งานแยกชิ้นส่วนและตัดแต่ง"]);

function getOTBreakMinutes(dept: string): number {
  return OT_LONG_BREAK_DEPTS.has(dept) ? 60 : 30;
}

function calcOTHoursForRecord(scanOut: string, shiftEnd: string, breakMinutes = 0): number {
  if (!scanOut || scanOut === "-" || !shiftEnd) return 0;
  if (!scanOut.includes(":") || !shiftEnd.includes(":")) return 0;
  const [soH, soM] = scanOut.split(":").map(Number);
  const [seH, seM] = shiftEnd.split(":").map(Number);
  const soTotal = soH * 60 + soM;
  // "00:00" from addHoursToTime wrapping past midnight means 24:00 (1440 min)
  const seTotal = (seH === 0 && seM === 0) ? 1440 : seH * 60 + seM;
  let diff = soTotal - seTotal;
  // Handle night-shift crossing midnight (e.g. shiftEnd=23:00, scanOut=01:30 → diff negative large)
  if (diff < -720) diff += 1440;
  // เวลาพักหลังเลิกงาน (30 นาที หรือ 1 ชม. แล้วแต่หน่วยงาน) ไม่นับเป็น OT
  diff -= breakMinutes;
  return Math.max(0, diff / 60);
}

type OTRecord = AttendanceRecord & { shiftEnd: string; otHours: number };

function OTDashboard({
  reportData,
  activeMasterMap,
  isLoadingReport,
  otSubTab,
  setOtSubTab,
  scanUploadedAt,
  holidayDates,
}: {
  reportData: ReportData | null;
  activeMasterMap: Partial<Record<MasterFileKey, MasterFile>>;
  isLoadingReport: boolean;
  otSubTab: "chart" | "summary" | "detail";
  setOtSubTab: Dispatch<SetStateAction<"chart" | "summary" | "detail">>;
  scanUploadedAt: string | null;
  holidayDates: Set<string>;
}) {
  const [otTarget, setOtTarget] = useState<number>(() => {
    if (typeof window === "undefined") return 2.0;
    const s = localStorage.getItem("ot_target");
    return s ? parseFloat(s) : 2.0;
  });
  const [deptManagers, setDeptManagers] = useState<Record<string, string>>(() => {
    if (typeof window === "undefined") return {};
    try { return JSON.parse(localStorage.getItem("ot_dept_managers") ?? "{}"); } catch { return {}; }
  });
  const [managerOptions, setManagerOptions] = useState<Array<{ empId: string; name: string; dept: string }>>([]);
  const [showConfig, setShowConfig] = useState(false);
  const [selectedDept, setSelectedDept] = useState<string | null>(null);
  const chartPanelRef = useRef<HTMLDivElement>(null);
  const [chartTrackH, setChartTrackH] = useState(200);

  const chartObserverRef = useRef<ResizeObserver | null>(null);
  const chartLayoutRef = (el: HTMLDivElement | null) => {
    if (chartObserverRef.current) { chartObserverRef.current.disconnect(); chartObserverRef.current = null; }
    if (!el) return;
    const obs = new ResizeObserver(([entry]) => {
      setChartTrackH(Math.max(150, Math.floor(entry.contentRect.height - 96)));
    });
    obs.observe(el);
    chartObserverRef.current = obs;
  };

  useEffect(() => {
    const path = activeMasterMap.dayoff_shift?.file_path;
    if (!path) return;
    let cancelled = false;
    (async () => {
      try {
        const rows = await downloadSheetRows(path);
        if (cancelled) return;
        const mgrs: Array<{ empId: string; name: string; dept: string }> = [];
        for (const row of rows) {
          const jobSite = String(row["หน่วยงานย่อย/Skill"] ?? row["หน้างาน"] ?? "").trim();
          if (!jobSite.includes("ผู้จัดการ")) continue;
          const empId = cleanEmpId(row["User ID (Job Information)"] ?? row["Employee ID"] ?? row["Emp ID"]);
          const firstName = String(row["First Name (Local)"] ?? "").trim();
          const lastName = String(row["Last Name (Local)"] ?? "").trim();
          const name = `${firstName} ${lastName}`.trim() || String(row["Employee Name"] ?? row["Name"] ?? "").trim();
          const dept = findRowCol(row, "หน่วยงาน", "Org. Unit Description", "Name (Section)", "แผนก", "Department");
          if (empId && name) mgrs.push({ empId, name, dept });
        }
        setManagerOptions(mgrs);
      } catch { /* master not yet loaded */ }
    })();
    return () => { cancelled = true; };
  }, [activeMasterMap.dayoff_shift?.file_path]);

  const otRecords = useMemo((): OTRecord[] => {
    if (!reportData) return [];
    return reportData.records.map((rec) => {
      const shiftEnd = rec.shiftEnd || (rec.shiftStart ? addHoursToTime(rec.shiftStart, 9) : "");
      const otHours = shiftEnd && (rec.status === "Present" || rec.status === "Late" || rec.status === "NoScanIn")
        ? calcOTHoursForRecord(rec.scanOut, shiftEnd, getOTBreakMinutes(rec.dept))
        : 0;
      return { ...rec, shiftEnd, otHours };
    });
  }, [reportData]);

  const isPublicHoliday = reportData?.isoTargetDate
    ? holidayDates.has(reportData.isoTargetDate)
    : false;

  const deptOTRows = useMemo(() => {
    const map = new Map<string, {
      dept: string; total: number; absent: number; late: number; dayoff: number;
      otWorkers: number; normalOTHours: number; publicHolidayOTHours: number;
      totalOTHours: number; activeWorkers: number;
    }>();
    for (const rec of otRecords) {
      const cur = map.get(rec.dept) ?? {
        dept: rec.dept, total: 0, absent: 0, late: 0, dayoff: 0,
        otWorkers: 0, normalOTHours: 0, publicHolidayOTHours: 0,
        totalOTHours: 0, activeWorkers: 0,
      };
      cur.total += 1;
      if (rec.status === "Absent") cur.absent += 1;
      if (rec.status === "Late") cur.late += 1;
      if (rec.status === "DayOff") cur.dayoff += 1;
      if (rec.status === "Present" || rec.status === "Late" || rec.status === "NoScanIn") {
        cur.activeWorkers += 1;
        if (rec.otHours > 0) {
          cur.otWorkers += 1;
          cur.totalOTHours += rec.otHours;
          if (isPublicHoliday) {
            cur.publicHolidayOTHours += rec.otHours;
          } else {
            cur.normalOTHours += rec.otHours;
          }
        }
      }
      map.set(rec.dept, cur);
    }
    return Array.from(map.values()).sort((a, b) => a.dept.localeCompare(b.dept, "th"));
  }, [otRecords, isPublicHoliday]);

  const totals = useMemo(() => ({
    total: deptOTRows.reduce((s, r) => s + r.total, 0),
    absent: deptOTRows.reduce((s, r) => s + r.absent, 0),
    late: deptOTRows.reduce((s, r) => s + r.late, 0),
    dayoff: deptOTRows.reduce((s, r) => s + r.dayoff, 0),
    otWorkers: deptOTRows.reduce((s, r) => s + r.otWorkers, 0),
    normalOTHours: deptOTRows.reduce((s, r) => s + r.normalOTHours, 0),
    publicHolidayOTHours: deptOTRows.reduce((s, r) => s + r.publicHolidayOTHours, 0),
    totalOTHours: deptOTRows.reduce((s, r) => s + r.totalOTHours, 0),
    activeWorkers: deptOTRows.reduce((s, r) => s + r.activeWorkers, 0),
  }), [deptOTRows]);

  const selectedDeptRecords = useMemo((): OTRecord[] => {
    if (!selectedDept) return otRecords;
    return otRecords.filter((r) => r.dept === selectedDept);
  }, [otRecords, selectedDept]);

  const [otDetailSearch, setOtDetailSearch] = useState("");
  const [otDetailDeptFilter, setOtDetailDeptFilter] = useState("all");
  const [otDetailSectionFilter, setOtDetailSectionFilter] = useState("all");
  const [otDetailShiftFilter, setOtDetailShiftFilter] = useState("all");

  const otDetailDeptOptions = useMemo(
    () => [...new Set(otRecords.map((r) => r.dept))].filter(Boolean).sort(),
    [otRecords],
  );
  const otDetailSectionOptions = useMemo(() => {
    const base = otDetailDeptFilter === "all" ? otRecords : otRecords.filter((r) => r.dept === otDetailDeptFilter);
    return [...new Set(base.map((r) => r.section))].filter(Boolean).sort();
  }, [otRecords, otDetailDeptFilter]);
  const otDetailShiftOptions = useMemo(() => {
    let base = selectedDept
      ? otRecords.filter((r) => r.dept === selectedDept)
      : otDetailDeptFilter !== "all"
        ? otRecords.filter((r) => r.dept === otDetailDeptFilter)
        : otRecords;
    if (otDetailSectionFilter !== "all") base = base.filter((r) => r.section === otDetailSectionFilter);
    return [...new Set(base.map((r) => r.shift))].filter(Boolean).sort();
  }, [otRecords, selectedDept, otDetailDeptFilter, otDetailSectionFilter]);

  const filteredDetailRecords = useMemo((): OTRecord[] => {
    const q = otDetailSearch.trim().toLowerCase();
    let base = selectedDept ? otRecords.filter((r) => r.dept === selectedDept) : otRecords;
    if (!selectedDept && otDetailDeptFilter !== "all") base = base.filter((r) => r.dept === otDetailDeptFilter);
    return base.filter((r) => {
      if (otDetailSectionFilter !== "all" && r.section !== otDetailSectionFilter) return false;
      if (otDetailShiftFilter !== "all" && r.shift !== otDetailShiftFilter) return false;
      if (q && !r.empId.toLowerCase().includes(q) && !r.name.toLowerCase().includes(q) && !r.dept.toLowerCase().includes(q)) return false;
      return true;
    });
  }, [otRecords, selectedDept, otDetailDeptFilter, otDetailSearch, otDetailSectionFilter, otDetailShiftFilter]);

  const totalAvgOT = totals.activeWorkers > 0 ? totals.totalOTHours / totals.activeWorkers : 0;
  const chartRows = [
    ...deptOTRows,
    { dept: "รวมทั้งหมด", activeWorkers: totals.activeWorkers, totalOTHours: totals.totalOTHours },
  ];
  const maxAvgOT = Math.max(
    ...chartRows.map((r) => (r.activeWorkers > 0 ? r.totalOTHours / r.activeWorkers : 0)),
    otTarget,
    0.1,
  );
  const yMax = Math.max(Math.ceil(maxAvgOT) + 1, 5);
  const yTicks = Array.from({ length: yMax + 1 }, (_, i) => i);
  const TRACK_H = chartTrackH;
  const XLAB_H = 96;
  const DEPT_BAR_COLORS = [
    "#ef4444", "#8b5cf6", "#3b82f6", "#10b981", "#f59e0b",
    "#ec4899", "#14b8a6", "#f97316", "#06b6d4", "#84cc16",
    "#a855f7", "#eab308", "#22c55e", "#0ea5e9", "#d946ef",
    "#64748b",
  ];

  function saveOTTarget(value: string) {
    const num = parseFloat(value);
    if (!isNaN(num) && num >= 0) {
      setOtTarget(num);
      localStorage.setItem("ot_target", String(num));
    }
  }

  function saveDeptManager(dept: string, name: string) {
    const next = { ...deptManagers, [dept]: name };
    setDeptManagers(next);
    localStorage.setItem("ot_dept_managers", JSON.stringify(next));
  }

  const now = new Date();
  const timeStr = `${String(now.getHours()).padStart(2, "0")}.${String(now.getMinutes()).padStart(2, "0")} น.`;

  return (
    <section className="ot-dashboard">
      {/* ── Header ── */}
      <div className="panel ot-header">
        <div className="ot-header-left">
          <h2 className="ot-title">การติดตาม OT ภายในหน่วยงาน (พนักงาน)</h2>
        </div>
        <div className="ot-header-right">
          <span className="ot-pull-badge">
            ดึงข้อมูล {scanUploadedAt
              ? (() => {
                  const d = new Date(scanUploadedAt);
                  const date = d.toLocaleDateString("th-TH", { day: "numeric", month: "short", year: "numeric" });
                  const time = d.toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit", hour12: false });
                  return date + "      " + time + " น.";
                })()
              : timeStr}
          </span>
          <button
            className={`ot-config-toggle${showConfig ? " active" : ""}`}
            onClick={() => setShowConfig((v) => !v)}
            type="button"
          >
            ⚙ ตั้งค่าเป้าหมาย
          </button>
        </div>
      </div>

      {/* ── Config Panel ── */}
      {showConfig && (
        <div className="panel ot-config">
          <div className="ot-config-row">
            <div className="ot-config-field">
              <label>เป้าหมาย OT (ชม./คน/วัน)</label>
              <input
                key={otTarget}
                type="number"
                min="0"
                step="0.5"
                defaultValue={otTarget}
                onBlur={(e) => saveOTTarget(e.target.value)}
                className="ot-cfg-input"
              />
            </div>
          </div>
        </div>
      )}

      {!reportData ? (
        <div className="panel ot-empty">
          {isLoadingReport ? (
            <p>กำลังโหลดข้อมูล...</p>
          ) : (
            <p>กรุณาเลือกวันที่และโหลดข้อมูลการเข้างานก่อนครับ</p>
          )}
        </div>
      ) : (
        <>
          {/* ── Bar Chart ── */}
          {otSubTab === "chart" && <div className="panel ot-chart-panel" ref={chartPanelRef}>
            <p className="ot-chart-title">
              วันที่ {reportData.targetDate}&nbsp;&nbsp;เปรียบเทียบ เฉลี่ยชั่วโมง O.T. ต่อคน/วัน
            </p>
            <div className="ot-chart-layout" ref={chartLayoutRef}>
              {/* Y-axis labels */}
              <div className="ot-yaxis-col" style={{ height: TRACK_H + XLAB_H }}>
                {yTicks.slice().reverse().map((tick) => (
                  <span
                    key={tick}
                    className="ot-ytick-label"
                    style={{ bottom: (tick / yMax) * TRACK_H + XLAB_H - 7 }}
                  >
                    {tick}
                  </span>
                ))}
              </div>
              {/* Chart track */}
              <div className="ot-chart-track-wrap" style={{ height: TRACK_H + XLAB_H }}>
                {/* Gridlines */}
                {yTicks.map((tick) => (
                  <div
                    key={tick}
                    className={`ot-gridline-h${tick === 0 ? " ot-gridline-base" : ""}`}
                    style={{ bottom: (tick / yMax) * TRACK_H + XLAB_H }}
                  />
                ))}
                {/* Target line */}
                <div
                  className="ot-target-line-h"
                  style={{ bottom: (otTarget / yMax) * TRACK_H + XLAB_H }}
                />
                {/* Bars */}
                <div className="ot-bars-flex" style={{ height: TRACK_H, bottom: XLAB_H }}>
                  {chartRows.map((row, i) => {
                    const avg = row.activeWorkers > 0 ? row.totalOTHours / row.activeWorkers : 0;
                    const barH = (avg / yMax) * TRACK_H;
                    const isTotal = row.dept === "รวมทั้งหมด";
                    const color = isTotal ? "#475569" : DEPT_BAR_COLORS[i % DEPT_BAR_COLORS.length];
                    return (
                      <div key={row.dept} className={`ot-bar-item-v2${isTotal ? " ot-bar-total-v2" : ""}`}>
                        <div className="ot-bar-val-v2">{avg.toFixed(1)}</div>
                        <div
                          className="ot-bar-fill-v2"
                          style={{ height: Math.max(barH, 0), backgroundColor: color }}
                        />
                      </div>
                    );
                  })}
                </div>
                {/* X-axis dept labels */}
                <div className="ot-xaxis-row" style={{ height: XLAB_H, bottom: 0 }}>
                  {chartRows.map((row) => (
                    <div key={row.dept} className="ot-xlab-item" title={row.dept}>
                      <span>{row.dept}</span>
                    </div>
                  ))}
                </div>
              </div>
              {/* Legend box */}
              <div className="ot-legend-box">
                <div className="ot-legend-box-inner">
                  <div className="ot-legend-arrow">⬇</div>
                  <strong>Target</strong>
                  <div>OT น้อยกว่า {otTarget.toFixed(1)} ชม.</div>
                  <div>ต่อคน ต่อวัน</div>
                </div>
              </div>
            </div>
          </div>}

          {/* ── Summary Table ── */}
          {otSubTab === "summary" && <div className="ot-table-scroll">
            <table className="table ot-table">
              <colgroup>
                <col style={{width: 130}} />{/* หน่วยงาน */}
                <col style={{width: 44}} />{/* อัตรา */}
                <col style={{width: 24}} /><col style={{width: 24}} /><col style={{width: 24}} />
                <col style={{width: 24}} />{/* ลาป่วย ลากิจ ลาบวช ลาพักผ่อน */}
                <col style={{width: 36}} />{/* ขาดงาน */}
                <col style={{width: 24}} />{/* ลาไม่จ่าย */}
                <col style={{width: 36}} />{/* รวมวันหยุด */}
                <col style={{width: 44}} />{/* %หยุด */}
                <col style={{width: 44}} />{/* ทำงาน */}
                <col style={{width: 44}} />{/* %OT */}
                <col style={{width: 44}} />{/* OT คน */}
                <col style={{width: 38}} />{/* เป้า */}
                <col style={{width: 46}} />{/* 1.5 */}
                <col style={{width: 36}} /><col style={{width: 36}} /><col style={{width: 36}} />{/* holiday ×3 */}
                <col style={{width: 50}} />{/* รวมชม. */}
                <col style={{width: 58}} />{/* เฉลี่ยคน/วัน */}
                <col style={{width: 108}} />{/* ผจก. */}
              </colgroup>
              <thead>
                <tr>
                  <th rowSpan={3} className="ot-th-sticky">
                    หน่วยงาน<br />
                    <span className="ot-th-date">{reportData.targetDate}</span>
                  </th>
                  <th rowSpan={3} className="ot-th-num">อัตรา<br />กำลังคน</th>
                  <th colSpan={7} className="ot-th-group">สถิติการมาทำงาน (คนงาน)</th>
                  <th rowSpan={3} className="ot-th-num">%<br />หยุดงาน</th>
                  <th rowSpan={3} className="ot-th-num ot-th-pink">พนักงาน<br />ที่ทำงาน</th>
                  <th rowSpan={3} className="ot-th-num ot-th-pink">%พนัก<br />ที่ OT</th>
                  <th rowSpan={3} className="ot-th-num ot-th-pink">พนักงาน<br />ที่ OT</th>
                  <th rowSpan={3} className="ot-th-num">เป้า<br />หมาย</th>
                  <th colSpan={5} className="ot-th-group ot-th-group-ot">เปรียบเทียบ ค่าล่วงเวลา (คนงาน)</th>
                  <th rowSpan={3} className="ot-th-avg-yellow">เฉลี่ย<br />คน/วัน</th>
                  <th rowSpan={3} className="ot-th-mgr ot-th-mgr-header">สถาพ / ผจก.</th>
                </tr>
                <tr>
                  <th rowSpan={2} className="ot-th-sub ot-th-vert">ลาป่วย</th>
                  <th rowSpan={2} className="ot-th-sub ot-th-vert">ลากิจ</th>
                  <th rowSpan={2} className="ot-th-sub ot-th-vert">ลาบวช<br />คลอด</th>
                  <th rowSpan={2} className="ot-th-sub ot-th-vert">ลาพัก<br />ผ่อน</th>
                  <th rowSpan={2} className="ot-th-sub ot-th-vert ot-th-absent">ขาดงาน</th>
                  <th rowSpan={2} className="ot-th-sub ot-th-vert">ลาไม่<br />จ่าย</th>
                  <th rowSpan={2} className="ot-th-sub ot-th-vert">รวมวัน<br />หยุดงาน</th>
                  <th className="ot-th-sub ot-th-ot-normal ot-th-subgrp">OT วันปกติ</th>
                  <th colSpan={3} className="ot-th-sub ot-th-holiday-grp ot-th-subgrp">OT วันหยุด</th>
                  <th rowSpan={2} className="ot-th-sub ot-th-vert-num">รวม<br />ชม.</th>
                </tr>
                <tr>
                  <th className="ot-th-sub ot-th-vert-num ot-th-ot-normal">1.5<br />ชม</th>
                  <th className="ot-th-sub ot-th-vert-num ot-th-holiday">1<br />ชม</th>
                  <th className="ot-th-sub ot-th-vert-num ot-th-holiday">2<br />ชม</th>
                  <th className="ot-th-sub ot-th-vert-num ot-th-holiday">3<br />ชม</th>
                </tr>
              </thead>
              <tbody>
                {deptOTRows.map((row) => {
                  const activeNonOff = row.total - row.dayoff;
                  const pctStop = activeNonOff > 0 ? Math.round((row.absent / activeNonOff) * 100) : 0;
                  const pctOT = row.activeWorkers > 0 ? Math.round((row.otWorkers / row.activeWorkers) * 100) : 0;
                  const avgOT = row.activeWorkers > 0 ? row.totalOTHours / row.activeWorkers : 0;
                  const mgrsForDept = managerOptions.filter((m) => m.dept === row.dept);
                  const mgrPool = mgrsForDept.length > 0 ? mgrsForDept : managerOptions;
                  return (
                    <tr
                      key={row.dept}
                      className={`ot-row${selectedDept === row.dept ? " ot-row-selected" : ""}`}
                      onClick={() => { const next = selectedDept === row.dept ? null : row.dept; setSelectedDept(next); setOtSubTab(next ? "detail" : "summary"); setOtDetailDeptFilter("all"); setOtDetailSectionFilter("all"); setOtDetailShiftFilter("all"); setOtDetailSearch(""); }}
                    >
                      <td className="ot-td-dept">{row.dept}</td>
                      <td className="ot-td-num">{row.total}</td>
                      <td className="ot-td-num ot-td-muted">-</td>
                      <td className="ot-td-num ot-td-muted">-</td>
                      <td className="ot-td-num ot-td-muted">-</td>
                      <td className="ot-td-num ot-td-muted">-</td>
                      <td className="ot-td-num ot-td-absent-cell">{row.absent > 0 ? row.absent : ""}</td>
                      <td className="ot-td-num ot-td-muted">-</td>
                      <td className="ot-td-num">{row.dayoff > 0 ? row.dayoff : ""}</td>
                      <td className="ot-td-num">{pctStop}%</td>
                      <td className="ot-td-num">{row.activeWorkers}</td>
                      <td className="ot-td-num">{row.activeWorkers > 0 ? `${pctOT}%` : ""}</td>
                      <td className={`ot-td-num${row.otWorkers > 0 ? " ot-hl-blue" : ""}`}>
                        {row.otWorkers > 0 ? row.otWorkers : ""}
                      </td>
                      <td className="ot-td-num">{otTarget.toFixed(1)}</td>
                      <td className="ot-td-num">{row.normalOTHours > 0 ? row.normalOTHours.toFixed(1) : "-"}</td>
                      <td className="ot-td-num ot-td-muted">-</td>
                      <td className={`ot-td-num${row.publicHolidayOTHours > 0 ? " ot-hl-orange" : " ot-td-muted"}`}>{row.publicHolidayOTHours > 0 ? row.publicHolidayOTHours.toFixed(1) : "-"}</td>
                      <td className="ot-td-num ot-td-muted">-</td>
                      <td className="ot-td-num">{row.totalOTHours > 0 ? row.totalOTHours.toFixed(1) : "-"}</td>
                      <td className={`ot-td-num ot-td-avg-yellow${avgOT > otTarget ? " ot-hl-red" : avgOT > 0 ? " ot-hl-green" : ""}`}>
                        {avgOT > 0 ? avgOT.toFixed(1) : "-"}
                      </td>
                      <td className="ot-td-mgr" onClick={(e) => e.stopPropagation()}>
                        <select
                          className="ot-mgr-select"
                          value={deptManagers[row.dept] ?? ""}
                          onChange={(e) => saveDeptManager(row.dept, e.target.value)}
                        >
                          <option value="">-</option>
                          {mgrPool.map((m) => (
                            <option key={m.empId} value={m.name}>{m.name}</option>
                          ))}
                        </select>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr className="ot-total-row">
                  <td>รวมทั้งหมด</td>
                  <td className="ot-td-num">{totals.total}</td>
                  <td className="ot-td-num ot-td-muted">-</td>
                  <td className="ot-td-num ot-td-muted">-</td>
                  <td className="ot-td-num ot-td-muted">-</td>
                  <td className="ot-td-num ot-td-muted">-</td>
                  <td className="ot-td-num ot-td-absent-cell">{totals.absent > 0 ? totals.absent : ""}</td>
                  <td className="ot-td-num ot-td-muted">-</td>
                  <td className="ot-td-num">{totals.dayoff > 0 ? totals.dayoff : ""}</td>
                  <td className="ot-td-num">
                    {Math.max(0, totals.total - totals.dayoff) > 0
                      ? `${Math.round((totals.absent / Math.max(0, totals.total - totals.dayoff)) * 100)}%`
                      : "0%"}
                  </td>
                  <td className="ot-td-num">{totals.activeWorkers}</td>
                  <td className="ot-td-num">
                    {totals.activeWorkers > 0
                      ? `${Math.round((totals.otWorkers / totals.activeWorkers) * 100)}%`
                      : ""}
                  </td>
                  <td className="ot-td-num">{totals.otWorkers > 0 ? totals.otWorkers : ""}</td>
                  <td className="ot-td-num">{otTarget.toFixed(1)}</td>
                  <td className="ot-td-num">{totals.normalOTHours > 0 ? totals.normalOTHours.toFixed(1) : "-"}</td>
                  <td className="ot-td-num ot-td-muted">-</td>
                  <td className={`ot-td-num${totals.publicHolidayOTHours > 0 ? " ot-hl-orange" : " ot-td-muted"}`}>{totals.publicHolidayOTHours > 0 ? totals.publicHolidayOTHours.toFixed(1) : "-"}</td>
                  <td className="ot-td-num ot-td-muted">-</td>
                  <td className="ot-td-num">{totals.totalOTHours > 0 ? totals.totalOTHours.toFixed(1) : "-"}</td>
                  <td className="ot-td-num ot-td-avg-yellow">{totalAvgOT > 0 ? totalAvgOT.toFixed(1) : "-"}</td>
                  <td />
                </tr>
              </tfoot>
            </table>
          </div>}

          {/* ── Employee Detail Table ── */}
          {otSubTab === "detail" && <div className="panel ot-detail-panel">
            <div className="ot-detail-hdr">
              <h3>
                รายละเอียดรายคน
                {selectedDept ? ` · ${selectedDept}` : ""}
                <span className="ot-detail-count"> ({filteredDetailRecords.length} คน)</span>
              </h3>
              <div className="ot-detail-filters">
                <input
                  type="text"
                  className="ot-detail-search"
                  placeholder="ค้นหา รหัส ชื่อ หน่วยงาน"
                  value={otDetailSearch}
                  onChange={(e) => setOtDetailSearch(e.target.value)}
                />
                <select value={otDetailDeptFilter} onChange={(e) => { setSelectedDept(null); setOtDetailDeptFilter(e.target.value); setOtDetailSectionFilter("all"); setOtDetailShiftFilter("all"); }}>
                  <option value="all">ทุกหน่วยงาน</option>
                  {otDetailDeptOptions.map((d) => <option key={d} value={d}>{d}</option>)}
                </select>
                <select style={{ width: 130 }} value={otDetailSectionFilter} onChange={(e) => { setOtDetailSectionFilter(e.target.value); setOtDetailShiftFilter("all"); }}>
                  <option value="all">ทุกหน่วยงานย่อย</option>
                  {otDetailSectionOptions.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
                <select value={otDetailShiftFilter} onChange={(e) => setOtDetailShiftFilter(e.target.value)}>
                  <option value="all">ทุก Shift</option>
                  {otDetailShiftOptions.map((s) => <option key={s} value={s}>{s}</option>)}
                </select>
                {selectedDept && (
                  <button type="button" className="ot-close-btn" onClick={() => {
                    setSelectedDept(null);
                    setOtSubTab("summary");
                    setOtDetailSectionFilter("all");
                    setOtDetailShiftFilter("all");
                    setOtDetailSearch("");
                  }}>
                    ✕ ล้างตัวกรอง
                  </button>
                )}
                <button
                  type="button"
                  className="primary-button small"
                  disabled={filteredDetailRecords.length === 0}
                  onClick={() => {
                    const rows = filteredDetailRecords.slice().sort((a, b) => b.otHours - a.otHours).map((r) => ({
                      "รหัส": r.empId,
                      "ชื่อ-นามสกุล": r.name,
                      "หน่วยงาน": r.dept,
                      "หน่วยงานย่อย": r.section,
                      "กะ": r.shift,
                      "เริ่มกะ": r.shiftStart,
                      "สิ้นสุดกะ": r.shiftEnd,
                      "สแกนเข้า": r.scanIn,
                      "สแกนออก": r.scanOut,
                      "สถานะ": STATUS_TH[r.status] ?? r.status,
                      "OT (ชม.)": r.otHours > 0 ? r.otHours.toFixed(2) : "",
                    }));
                    const ws = XLSX.utils.json_to_sheet(rows);
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, "OT Detail");
                    XLSX.writeFile(wb, `OT-รายละเอียดรายคน.xlsx`);
                  }}
                >
                  <Download size={14} />
                  Export Excel
                </button>
              </div>
            </div>
            <div className="table-scroll">
              <table className="table">
                <thead>
                  <tr>
                    <th>รหัส</th>
                    <th>ชื่อ-นามสกุล</th>
                    <th>หน่วยงาน</th>
                    <th>กะ</th>
                    <th>เริ่มกะ</th>
                    <th>สิ้นสุดกะ</th>
                    <th>สแกนเข้า</th>
                    <th>สแกนออก</th>
                    <th>สถานะ</th>
                    <th>OT (ชม.)</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredDetailRecords
                    .slice()
                    .sort((a, b) => b.otHours - a.otHours)
                    .map((rec, i) => (
                      <tr key={rec.empId + "-" + i} className={rec.otHours > 0 ? "ot-detail-has-ot" : ""}>
                        <td>{rec.empId}</td>
                        <td>{rec.name}</td>
                        <td>{rec.dept}</td>
                        <td>{rec.shift}</td>
                        <td>{rec.shiftStart}</td>
                        <td>{rec.shiftEnd}</td>
                        <td className="scan-cell">{scanDateBadge(rec.scanInDate)}{rec.scanIn}</td>
                        <td className="scan-cell">{scanDateBadge(rec.scanOutDate)}{rec.scanOut}</td>
                        <td>
                          <span className={`ot-status-badge ot-status-${rec.status.toLowerCase()}`}>
                            {STATUS_TH[rec.status] ?? rec.status}
                          </span>
                        </td>
                        <td className={rec.otHours > 0 ? "ot-td-positive" : "ot-td-zero"}>
                          {rec.otHours > 0 ? rec.otHours.toFixed(2) : "-"}
                        </td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
          </div>}
        </>
      )}
    </section>
  );
}

type HelpSection = {
  id: string;
  title: string;
  group: string;
  icon: typeof HomeIcon;
  summary: string;
  points: string[];
  openTab?: TabId;
  openMasterSubTab?: MasterSubTab;
  openOtSubTab?: OtSubTab;
};

const helpGroupOrder = ["ภาพรวม", "หน้าหลัก", "Master Data", "OT Dashboard"];

const helpSections: HelpSection[] = [
  {
    id: "getting-started",
    title: "เริ่มต้นใช้งาน — ลำดับที่ควรทำ",
    group: "ภาพรวม",
    icon: ClipboardCheck,
    summary: "ระบบต้องมีข้อมูลตั้งต้นให้ครบก่อน ถึงจะใช้งานหน้าอื่นๆ ได้เต็มที่ ทำตามลำดับนี้ในการใช้งานครั้งแรก",
    points: [
      "1) ไปที่ Master Data → Master Files แล้วอัปโหลดไฟล์รายชื่อพนักงาน (ต้องมีให้ครบ 4 ชนิด: รายชื่อพนักงาน, Manpower Plan, Skill Matrix, Dayoff & Shift) ระบบจะยังสร้าง Daily Run ไม่ได้จนกว่าจะมีไฟล์ทั้ง 4 ชนิดนี้ active อยู่",
      "2) ไปที่ Master Data → Manpower ตั้งเวลาเข้า-ออกของแต่ละหน่วยงาน+หน่วยงานย่อย+กะ ให้ครบ เพราะหน้า Shift & Dayoff จะดึงเวลานี้ไปใช้อัตโนมัติ",
      "3) ไปที่ Master Data → Shift & Dayoff กำหนดวันหยุดประจำสัปดาห์และกะทำงานของพนักงานแต่ละคน",
      "4) ไปที่ Upload Timestamp อัปโหลดไฟล์สแกนเข้า-ออกประจำวัน (.csv/.xlsx/.xls) แล้วกด Create Daily Run",
      "5) หลังจากนั้น Dashboard, ผลลัพธ์การจัดสรร, Timestamp With Dept, Report & Dashboard และ OT Dashboard จะมีข้อมูลให้ดูตามวันที่เลือกจากปุ่มปฏิทิน มุมขวาบนของทุกหน้า",
      "เคล็ดลับ: ปุ่มปฏิทิน (ข้อมูลวันที่) ที่มุมขวาบนของทุกหน้า ใช้เลือกดูข้อมูลของวันอื่นที่เคยอัปโหลดไว้แล้ว — จะเลือกได้เฉพาะวันที่มีการอัปโหลด timestamp ไว้แล้วเท่านั้น",
    ],
  },
  {
    id: "dashboard",
    title: "Dashboard",
    group: "หน้าหลัก",
    icon: HomeIcon,
    openTab: "dashboard",
    summary: "หน้าสรุปภาพรวมการเข้างานของวันที่เลือก ใช้ดูสถานะรายวันและยืนยันตรวจสอบการเข้างาน",
    points: [
      "การ์ดตัวเลข 5 ใบด้านบน (มาทำงาน / ตรงเวลา / มาสาย / ขาด-ลา / วันหยุด) กดได้ — จะเลื่อนลงไปที่ตารางรายละเอียดและกรองให้อัตโนมัติ",
      "กล่อง \"ยืนยันตรวจสอบการเข้างาน\": หัวหน้างานพิมพ์ชื่อแล้วกด \"ยืนยันรับทราบ\" เพื่อบันทึกว่าตรวจสอบข้อมูลวันนั้นแล้ว — ถ้าไม่กรอกชื่อ ระบบจะเตือนไม่ให้บันทึก กดปุ่ม \"แก้ไข\" เพื่อยกเลิกและยืนยันใหม่ได้",
      "ตารางคนมาสาย มีปุ่ม \"เตือน\" ต่อคน สำหรับบันทึกว่าได้เตือนพนักงานคนนั้นแล้ว และมี badge สีแดง/เหลืองแสดงความเสี่ยง (มาสายตั้งแต่ 3 ครั้งขึ้นไปในเดือนนี้ถือว่าเสี่ยง)",
      "แถวคนขาดงาน มีดรอปดาวน์เลือกประเภทการลา (ลาป่วย/ลากิจ/ลาพักร้อน ฯลฯ) ให้เลือกบันทึกได้ทันที",
      "ปุ่ม \"Export สาย/ขาด\" ที่มุมตาราง ใช้ export เฉพาะรายชื่อคนมาสาย/ขาดงานเป็นไฟล์ Excel",
      "กล่อง \"สถานะการตักเตือน\" (พับเก็บได้) แสดง % ความคืบหน้าการเตือนคนมาสายแยกตามหน่วยงาน สำหรับหัวหน้างานใช้ติดตามงาน",
    ],
  },
  {
    id: "timestamp",
    title: "Upload Timestamp",
    group: "หน้าหลัก",
    icon: UploadCloud,
    openTab: "timestamp",
    summary: "หน้าอัปโหลดไฟล์สแกนเข้า-ออกประจำวัน เพื่อสร้าง Daily Run ที่หน้าอื่นๆ ใช้อ้างอิงข้อมูล",
    points: [
      "ลากไฟล์วาง หรือคลิกเลือกไฟล์ (.csv, .xlsx, .xls) แล้วกด \"Create Daily Run\"",
      "ต้องมี Master Data ครบทั้ง 4 ไฟล์ (active อยู่) ก่อน ปุ่ม Create Daily Run ถึงจะกดได้ ไม่งั้นระบบจะเตือนว่า \"กรุณา upload master files ให้ครบก่อนสร้าง daily run\"",
      "ประวัติการอัปโหลดด้านล่างมีปุ่มดาวน์โหลด/ลบไฟล์แต่ละรายการ (การลบยืนยันซ้ำก่อนลบจริง เพราะย้อนกลับไม่ได้)",
      "มีแถบแสดงพื้นที่จัดเก็บที่ใช้ไป — ถ้าใกล้เต็ม (ราว 1001MB จาก 1024MB) ระบบจะเตือนให้ลบไฟล์เก่าก่อนอัปโหลดไฟล์ใหม่ พร้อมปุ่มลบไฟล์เก่าสุดให้ทันที",
    ],
  },
  {
    id: "results",
    title: "ผลลัพธ์การจัดสรร",
    group: "หน้าหลัก",
    icon: BriefcaseBusiness,
    openTab: "results",
    summary: "ตารางรายชื่อพนักงานที่มาทำงาน (ตรงเวลา/มาสาย) ของวันที่เลือก แบบแบ่งหน้า",
    points: [
      "ค้นหา, กรองตามหน่วยงาน, กรองตามสถานะ (เฉพาะตรงเวลา/มาสาย ไม่รวมขาดงาน) และปุ่ม Clear ล้างตัวกรอง",
      "แสดงคอลัมน์ รหัสพนักงาน, ชื่อ, หน่วยงาน, ตำแหน่ง, หน่วยงานย่อย, กะ, เวลาสแกนเข้า, สถานะ",
      "หมายเหตุ: หน้านี้ยังไม่มีปุ่ม Export ที่ใช้งานได้จริง ถ้าต้องการ export ข้อมูล แนะนำใช้หน้า Timestamp With Dept แทน ซึ่งมีข้อมูลครบกว่า (รวมคนขาดงาน/วันหยุดด้วย) และปุ่ม Export ใช้งานได้จริง",
    ],
  },
  {
    id: "timestamp-dept",
    title: "Timestamp With Dept",
    group: "หน้าหลัก",
    icon: Database,
    openTab: "timestamp_dept",
    summary: "ตารางข้อมูลการสแกนเข้า-ออกแบบเต็ม รวมทุกสถานะ (มาทำงาน/มาสาย/ขาดงาน/วันหยุด/ขาดสแกนเข้า/รอเข้างาน) ผูกกับหน่วยงานและกะ",
    points: [
      "ค้นหา, กรองตามหน่วยงาน, กรองตามสถานะแบบละเอียด, เรียงลำดับได้ทุกคอลัมน์ (กดหัวตาราง), แบ่งหน้า 10 แถว/หน้า",
      "ปุ่ม \"Export Excel\" ใช้งานได้จริง — export ตามตัวกรอง/การเรียงลำดับที่ตั้งไว้ รวมเวลาสแกนเข้า-ออก และจำนวนนาทีที่มาสาย",
      "ถ้ายังไม่เคยเลือกวันที่/โหลดข้อมูล จะขึ้นข้อความว่ายังไม่มีข้อมูล — ต้องเลือกวันที่จากปุ่มปฏิทินก่อน",
    ],
  },
  {
    id: "master-files",
    title: "Master Data → Master Files",
    group: "Master Data",
    icon: FileSpreadsheet,
    openTab: "master",
    openMasterSubTab: "files",
    summary: "จัดการไฟล์รายชื่อพนักงานรวม (Employee Master + Dayoff/Shift + Skill) และไฟล์ Manpower Plan",
    points: [
      "การ์ดซ้าย (ไฟล์รวมพนักงาน): \"เทมเพลตเปล่า\" ดาวน์โหลดไฟล์ตัวอย่างเปล่า, \"Export ข้อมูลปัจจุบัน\" ดึงข้อมูลปัจจุบันทั้งหมดออกมาเป็นไฟล์เดียว",
      "อัปโหลดไฟล์ใหม่ (.xlsx/.xls ไม่เกิน 50MB) แล้วกด \"ตรวจสอบก่อนอัพโหลด\" — ระบบจะเปิดหน้าต่างเทียบความแตกต่าง (เพิ่ม/ลบ/แก้ไข/ไม่เปลี่ยน) ให้ตรวจสอบก่อนกดยืนยันจริง",
      "ถ้ามีพนักงานที่จะถูกลบออกจากระบบ ต้องติ๊กยืนยันก่อนถึงจะบันทึกได้ (กันการลบพลาด)",
      "ถ้ามีรหัสพนักงานซ้ำในไฟล์ที่อัปโหลด ระบบจะไม่ให้บันทึกและแจ้งรหัสที่ซ้ำให้ทราบ",
      "การ์ดขวา (Manpower Plan): อัปโหลดไฟล์ (.xlsx/.xls/.csv) แล้วกด \"Save Manpower Plan\" — การอัปโหลดนี้ไม่กระทบรายชื่อพนักงาน แก้เฉพาะเวลากะ",
      "ทั้งสองการ์ดมีประวัติไฟล์ ดาวน์โหลด/ลบย้อนหลังได้ และมีแถบแสดงพื้นที่จัดเก็บที่ใช้ไป",
    ],
  },
  {
    id: "manpower",
    title: "Master Data → Manpower",
    group: "Master Data",
    icon: BarChart3,
    openTab: "master",
    openMasterSubTab: "manpower",
    summary: "ตารางกำหนดเวลาเข้า-ออกงานของแต่ละหน่วยงาน + หน่วยงานย่อย + กะ ซึ่งเป็นต้นทางเวลาที่หน้า Shift & Dayoff ใช้อัตโนมัติ",
    points: [
      "กรองตามหน่วยงาน/หน่วยงานย่อย/กะ และเรียงลำดับได้ทุกคอลัมน์ (กด 1 ครั้งเรียงมากไปน้อย กดอีกครั้งเรียงน้อยไปมาก)",
      "เพิ่มแถวใหม่: กรอกหน่วยงาน (จำเป็น), หน่วยงานย่อย (เว้นว่างได้ = ใช้กับทั้งหน่วยงาน), ชื่อกะ, เวลาเข้า-ออก แล้วกด \"+ เพิ่มหน่วยงาน/กะ\" — ถ้าไม่กรอกเวลาออก ระบบจะตั้งให้อัตโนมัติเป็นเวลาเข้า+9 ชั่วโมง",
      "แก้เวลาเข้า-ออกของแถวที่มีอยู่ได้โดยตรงในตาราง, ลบแถวด้วยไอคอนถังขยะ",
      "ชื่อกะ \"ผู้จัดการ\" เป็นคำสงวน หมายถึง \"ไม่มีกะตายตัว\" ห้ามตั้งเป็นชื่อกะจริง ระบบจะเตือนถ้าพยายามเพิ่ม",
      "ถ้ามีหน่วยงาน+หน่วยงานย่อย+กะซ้ำกัน ระบบจะเตือนสีแดงและไม่ให้บันทึกจนกว่าจะลบแถวที่ซ้ำออก",
      "อย่าลืมกด \"Save\" หลังแก้ไข ไม่งั้นการเปลี่ยนแปลงจะไม่ถูกบันทึก — ปุ่ม Save จะโชว์ \"(มีการแก้ไข)\" เตือนถ้ายังไม่ได้บันทึก",
    ],
  },
  {
    id: "holy-days",
    title: "Master Data → วันพระ",
    group: "Master Data",
    icon: CalendarDays,
    openTab: "master",
    openMasterSubTab: "holidays",
    summary: "จัดการวันหยุดตามปฏิทินพุทธศาสนา (วันพระ) ที่ใช้ประกอบการคำนวณวันหยุด/กะของพนักงาน",
    points: [
      "เพิ่มวันหยุด: เลือกวันที่ + ชื่อ + ประเภท (วันพระ/วันหยุดราชการ/วันหยุดบริษัท) แล้วกด \"+ เพิ่มวันหยุด\"",
      "มีแท็บปีให้เลือกดูย้อนหลัง/ล่วงหน้า และลบรายการด้วยปุ่ม × ที่แต่ละแถว",
      "ถ้าเพิ่มวันที่ซ้ำกับที่มีอยู่แล้ว ระบบจะเตือนให้ลบตัวเดิมก่อน",
    ],
  },
  {
    id: "public-holidays",
    title: "Master Data → วันหยุดประจำปี",
    group: "Master Data",
    icon: CalendarOff,
    openTab: "master",
    openMasterSubTab: "public_holidays",
    summary: "จัดการวันหยุดราชการ/วันหยุดบริษัทประจำปี ซึ่งมีผลต่อการคำนวณ OT วันหยุดในหน้า OT Dashboard",
    points: [
      "ระบบจะเติมวันหยุดราชการที่รู้จักให้อัตโนมัติในการเปิดหน้าครั้งแรก (เติมเฉพาะวันที่ยังไม่มีในระบบ ไม่ทับข้อมูลเดิม)",
      "เพิ่ม/ลบวันหยุดได้เหมือนหน้าวันพระ และมีคอลัมน์หมายเหตุแจ้งถ้าวันนั้นตรงกับวันพระด้วย",
      "สำคัญ: วันหยุดที่ตั้งไว้ในหน้านี้และหน้าวันพระ มีผลกับหน้า OT Dashboard โดยตรง — ถ้าวันที่เลือกอยู่ในรายการนี้ ชั่วโมง OT ของวันนั้นจะถูกนับเป็น \"OT วันหยุด\" แทน \"OT วันปกติ\"",
    ],
  },
  {
    id: "dayoff-shift",
    title: "Master Data → Shift & Dayoff",
    group: "Master Data",
    icon: CalendarClock,
    openTab: "master",
    openMasterSubTab: "dayoff_shift",
    summary: "หน้าหลักสำหรับกำหนดวันหยุดประจำสัปดาห์และกะทำงานของพนักงานแต่ละคน ซึ่งเป็นข้อมูลจริงที่ระบบใช้คำนวณการมาสาย/ขาดงาน",
    points: [
      "ค้นหา/กรองตามหน่วยงาน, หน่วยงานย่อย, วันหยุด, กะ",
      "แก้ไขทีละคนได้โดยตรงในตาราง (เลือกหน่วยงานย่อย, วันหยุด, กะ จากดรอปดาวน์)",
      "แก้ทีละหลายคน: ติ๊กเลือกพนักงานที่ต้องการ แล้วใช้กล่องสีฟ้าด้านบนเปลี่ยน Dayoff / หน่วยงานย่อย / Shift พร้อมกันได้ กด Apply เพื่อใช้กับทุกคนที่เลือก",
      "ปุ่ม \"⇄ สลับกะ 1 ↔ 2\" (โชว์เมื่อกรองหน่วยงาน+หน่วยงานย่อยเจาะจงแล้ว) ใช้สลับพนักงานระหว่างกะ1กับกะ2 ทั้งกลุ่มในคลิกเดียว",
      "ปุ่ม \"แนะนำหน่วยงานย่อยจาก Manpower\" ช่วยเทียบและแก้หน่วยงานย่อยของพนักงานให้ตรงกับ Manpower Plan อัตโนมัติ (เฉพาะกรณีที่หน่วยงาน+กะนั้นมีหน่วยงานย่อยที่ชัดเจนไม่ก้ำกึ่งเท่านั้น)",
      "สำคัญ: เวลาเข้า-ออกงานไม่ได้แก้ตรงนี้ — ระบบดึงมาจากหน้า Manpower อัตโนมัติตามหน่วยงาน+หน่วยงานย่อย+กะ ถ้าไม่พบเวลาที่ตรงกัน ช่องเวลาจะโชว์ \"—\" สีเหลือง แปลว่าไปตั้งค่าที่หน้า Manpower ก่อน",
      "ก่อนกด Save ถ้ามีพนักงานที่ไม่พบเวลาใน Manpower ระบบจะเตือนให้ทราบก่อนว่าเวลาของคนกลุ่มนั้นจะถูกบันทึกเป็นค่าว่าง",
    ],
  },
  {
    id: "skill-matrix",
    title: "Skill Matrix",
    group: "หน้าหลัก",
    icon: LayoutGrid,
    openTab: "skill",
    summary: "หน้าจัดการระดับความชำนาญ (Skill) ของพนักงานแต่ละคนในแต่ละทักษะ",
    points: [
      "ค้นหา/กรองตามหน่วยงาน, ทักษะ, กะ",
      "แก้ทีละคนได้จากดรอปดาวน์ Level ในตาราง (0=ยังไม่ระบุ, 1=น้อย, 2=ปานกลาง, 3=ถนัด)",
      "แก้ทีละหลายคน: ติ๊กเลือกพนักงาน แล้วใช้กล่อง \"เปลี่ยน Level...\" ด้านบน กด Apply",
      "ปุ่ม \"+ เพิ่ม Skill ให้พนักงาน\" เปิดฟอร์มเพิ่มคู่พนักงาน+ทักษะใหม่ พิมพ์รหัสพนักงาน/ชื่อทักษะจะมีตัวช่วยเดาให้ (autocomplete)",
      "ตอนกด Save ระบบจะบันทึกเฉพาะรายการที่มี Level มากกว่า 0 เท่านั้น รายการที่ยังไม่ระบุ (Level 0) จะไม่ถูกบันทึก",
    ],
  },
  {
    id: "report-dashboard",
    title: "Report & Dashboard",
    group: "หน้าหลัก",
    icon: BarChart3,
    openTab: "report",
    summary: "หน้าวิเคราะห์เชิงลึกของการเข้างานในวันที่เลือก แยกดูตามหน่วยงานได้ พร้อมสถานะการตรวจสอบของหัวหน้างาน",
    points: [
      "เลือกหน่วยงานจากดรอปดาวน์ แล้วกด \"โหลดข้อมูล\" เพื่อดึงข้อมูลล่าสุด",
      "กราฟแท่ง \"การเข้างานรายแผนก\" กดที่แท่งไหนได้ จะเจาะลึกไปดูเฉพาะหน่วยงานนั้นทั้งหน้า (มีปุ่ม \"ดูทั้งหมด\" ไว้ยกเลิกการเจาะจง)",
      "กล่องพับเก็บได้ 2 อัน: \"สถานะการตักเตือน\" (ความคืบหน้าการเตือนคนมาสายต่อหน่วยงาน) และ \"สถานะการขาด/ลา\" (ความคืบหน้าการระบุประเภทลา)",
      "ตาราง \"รายละเอียด Late & Absent\" ค้นหา/กรอง/เรียงลำดับได้ และดูประวัติการเตือนย้อนหลังต่อคนได้",
      "ตาราง \"สถานะการตรวจสอบการเข้างานรายหน่วยงาน\" สรุปว่าหน่วยงานไหนหัวหน้างานยืนยันตรวจสอบแล้วบ้าง (ผูกกับปุ่มยืนยันในหน้า Dashboard)",
      "ถ้ายังไม่มีไฟล์รายชื่อพนักงานหรือยังไม่มี timestamp ล่าสุด ระบบจะแจ้งเตือนให้อัปโหลดให้ครบก่อน",
    ],
  },
  {
    id: "ot-chart",
    title: "OT Dashboard → แผนภูมิ",
    group: "OT Dashboard",
    icon: TrendingUp,
    openTab: "ot",
    openOtSubTab: "chart",
    summary: "กราฟแท่งเปรียบเทียบชั่วโมง OT เฉลี่ยต่อคนต่อวัน แยกตามหน่วยงาน เทียบกับเป้าหมายที่ตั้งไว้",
    points: [
      "มีเส้นเป้าหมาย (target) ในกราฟ ปรับได้จากปุ่ม \"⚙ ตั้งค่าเป้าหมาย\" ที่มุมบน",
      "ต้องเลือกวันที่และโหลดข้อมูลการเข้างานก่อน ถึงจะเห็นกราฟ",
    ],
  },
  {
    id: "ot-summary",
    title: "OT Dashboard → สรุปรายหน่วยงาน",
    group: "OT Dashboard",
    icon: LayoutGrid,
    openTab: "ot",
    openOtSubTab: "summary",
    summary: "ตารางสรุป OT แยกตามหน่วยงาน (จำนวนคน, ขาด/ลา, % OT, ชั่วโมง OT ปกติ/วันหยุด, ผู้จัดการที่รับผิดชอบ)",
    points: [
      "กดที่แถวหน่วยงานไหน จะพาไปหน้า \"สรุปรายพนักงาน\" พร้อมกรองเฉพาะหน่วยงานนั้นให้อัตโนมัติ",
      "ดรอปดาวน์ \"ผจก.\" ในแต่ละแถว เลือกผู้จัดการที่ดูแลหน่วยงานนั้นได้ (ดึงรายชื่อจากคนที่มีคำว่า \"ผู้จัดการ\" ในหน่วยงานย่อย)",
      "ค่าเป้าหมาย OT และผู้จัดการที่เลือกไว้ บันทึกเก็บไว้ในเบราว์เซอร์เครื่องนี้เท่านั้น (localStorage) ไม่ได้บันทึกลงฐานข้อมูลกลาง จึงไม่แชร์ข้ามเครื่อง/ผู้ใช้คนอื่น",
    ],
  },
  {
    id: "ot-detail",
    title: "OT Dashboard → สรุปรายพนักงาน",
    group: "OT Dashboard",
    icon: UsersRound,
    openTab: "ot",
    openOtSubTab: "detail",
    summary: "ตารางชั่วโมง OT รายบุคคล เรียงจากมากไปน้อย พร้อมตัวกรองและ Export",
    points: [
      "ค้นหา/กรองตามหน่วยงาน, หน่วยงานย่อย, กะ",
      "ปุ่ม \"Export Excel\" ใช้งานได้จริง export ตามตัวกรองที่ตั้งไว้",
      "ปุ่ม \"✕ ล้างตัวกรอง\" จะปรากฏเมื่อมาจากการกดเจาะจงหน่วยงานจากหน้าอื่น ใช้ล้างเพื่อดูข้อมูลทุกหน่วยงานอีกครั้ง",
      "ชั่วโมง OT คำนวณจาก (เวลาสแกนออก − เวลาเลิกกะ) เฉพาะคนที่มาทำงาน/มาสาย/ขาดสแกนเข้าเท่านั้น",
    ],
  },
  {
    id: "setting",
    title: "Setting",
    group: "หน้าหลัก",
    icon: Settings,
    openTab: "setting",
    summary: "หน้าตั้งค่าระบบ — ขณะนี้ยังไม่เปิดใช้งาน",
    points: [
      "หน้านี้ยังอยู่ระหว่างพัฒนา ยังไม่มีตัวเลือกให้ตั้งค่าจริง",
    ],
  },
];

function highlightMatches(text: string, query: string): ReactNode {
  if (!query) return text;
  const lower = text.toLowerCase();
  const q = query.toLowerCase();
  const parts: ReactNode[] = [];
  let start = 0;
  let idx = lower.indexOf(q, start);
  while (idx !== -1) {
    if (idx > start) parts.push(text.slice(start, idx));
    parts.push(<mark key={idx}>{text.slice(idx, idx + query.length)}</mark>);
    start = idx + query.length;
    idx = lower.indexOf(q, start);
  }
  parts.push(text.slice(start));
  return parts;
}

const helpNotePrefixRe = /^(สำคัญ|หมายเหตุ|เคล็ดลับ):\s*/;

function splitPointsAndNotes(points: string[]): { steps: string[]; notes: string[] } {
  const steps: string[] = [];
  const notes: string[] = [];
  for (const p of points) {
    (helpNotePrefixRe.test(p) ? notes : steps).push(p);
  }
  return { steps, notes };
}

function HelpGuidePage({
  setActiveTab,
  setMasterSubTab,
  setOtSubTab,
}: {
  setActiveTab: (tab: TabId) => void;
  setMasterSubTab: (tab: MasterSubTab) => void;
  setOtSubTab: (tab: OtSubTab) => void;
}) {
  const [query, setQuery] = useState("");
  const [activeId, setActiveId] = useState(helpSections[0]?.id ?? "");
  const q = query.trim().toLowerCase();

  function openSection(s: HelpSection) {
    if (!s.openTab) return;
    setActiveTab(s.openTab);
    if (s.openMasterSubTab) setMasterSubTab(s.openMasterSubTab);
    if (s.openOtSubTab) setOtSubTab(s.openOtSubTab);
  }

  const filtered = q
    ? helpSections.filter(
        (s) =>
          s.title.toLowerCase().includes(q) ||
          s.summary.toLowerCase().includes(q) ||
          s.points.some((p) => p.toLowerCase().includes(q)),
      )
    : helpSections;

  const groups = helpGroupOrder
    .map((group) => ({ group, sections: filtered.filter((s) => s.group === group) }))
    .filter((g) => g.sections.length > 0);

  return (
    <section className="panel help-guide-panel">
      <div className="help-guide-header">
        <div className="help-guide-title">
          <span className="help-guide-icon"><BookOpen size={22} /></span>
          <div>
            <h2>คู่มือการใช้งานระบบ</h2>
            <p>คำอธิบายแต่ละหน้าในระบบ สำหรับผู้ที่เพิ่งเริ่มใช้งาน</p>
          </div>
        </div>
        <div className="help-search-wrap">
          <Search size={16} />
          <input
            type="search"
            aria-label="ค้นหาคู่มือการใช้งาน"
            placeholder="ค้นหาคำสำคัญ เช่น อัปโหลด, กะ, Manpower, Export..."
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
          {query ? (
            <button type="button" className="help-search-clear" onClick={() => setQuery("")} aria-label="ล้างคำค้นหา">
              <X size={14} />
            </button>
          ) : null}
        </div>
      </div>

      {q ? (
        <p className="help-search-count">
          พบ {filtered.length} หัวข้อที่ตรงกับ &quot;{query}&quot;
        </p>
      ) : null}

      {filtered.length === 0 ? (
        <p className="empty-copy">ไม่พบหัวข้อที่ตรงกับคำค้นหา ลองใช้คำอื่นดูครับ</p>
      ) : (
        <div className="help-guide-body">
          <nav className="help-guide-sidebar" aria-label="สารบัญคู่มือ">
            {groups.map(({ group, sections }) => (
              <div key={group} className="help-sidebar-group">
                <span className="help-sidebar-group-label">{group}</span>
                {sections.map((s) => {
                  const SectionIcon = s.icon;
                  return (
                    <a
                      key={s.id}
                      href={`#help-${s.id}`}
                      className={`help-sidebar-link${activeId === s.id ? " active" : ""}`}
                      onClick={() => setActiveId(s.id)}
                    >
                      <SectionIcon size={15} />
                      <span>{highlightMatches(s.title, q)}</span>
                    </a>
                  );
                })}
              </div>
            ))}
          </nav>

          <div className="help-guide-content">
            {groups.map(({ group, sections }) => (
              <div key={group} className="help-content-group">
                <h3 className="help-content-group-title">{group}</h3>
                <div className="help-guide-sections">
                  {sections.map((s) => {
                    const SectionIcon = s.icon;
                    const { steps, notes } = splitPointsAndNotes(s.points);
                    return (
                      <article key={s.id} id={`help-${s.id}`} className="help-section-card">
                        <div className="help-section-hdr">
                          <span className="help-section-icon"><SectionIcon size={18} /></span>
                          <h3>{highlightMatches(s.title, q)}</h3>
                          {s.openTab ? (
                            <button type="button" className="help-open-btn" onClick={() => openSection(s)}>
                              เปิดหน้า <ArrowRight size={13} />
                            </button>
                          ) : null}
                        </div>
                        <p className="help-section-summary">{highlightMatches(s.summary, q)}</p>
                        <ul className="help-section-points">
                          {steps.map((p, i) => (
                            <li key={i}>{highlightMatches(p, q)}</li>
                          ))}
                        </ul>
                        {notes.length > 0 ? (
                          <div className="help-section-notes">
                            <span className="help-notes-label">เพิ่มเติม</span>
                            <ul>
                              {notes.map((p, i) => (
                                <li key={i}>{highlightMatches(p.replace(helpNotePrefixRe, ""), q)}</li>
                              ))}
                            </ul>
                          </div>
                        ) : null}
                      </article>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </section>
  );
}
