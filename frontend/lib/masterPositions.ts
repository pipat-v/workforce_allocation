import * as XLSX from "xlsx";
import { supabase } from "@/lib/supabase";

export async function loadActiveMasterPositions(): Promise<string[]> {
  const { data: master, error: masterError } = await supabase
    .from("master_data_files")
    .select("file_path")
    .eq("file_type", "employee_master")
    .eq("is_active", true)
    .order("created_at", { ascending: false })
    .limit(1)
    .maybeSingle();
  if (masterError) throw new Error(masterError.message);
  if (!master?.file_path) throw new Error("ยังไม่มี Employee Master ที่ Active");

  const { data: file, error: downloadError } = await supabase.storage
    .from("workforce-inputs")
    .download(master.file_path);
  if (downloadError || !file) {
    throw new Error(downloadError?.message ?? "ดาวน์โหลด Employee Master ไม่สำเร็จ");
  }

  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
  const positions = new Set<string>();
  for (const row of rows) {
    const positionKey = Object.keys(row).find((key) => {
      const normalized = key.trim().toLocaleLowerCase("th-TH");
      return normalized === "title (position)" || normalized === "position" || normalized === "ตำแหน่ง";
    });
    const position = positionKey ? String(row[positionKey] ?? "").trim() : "";
    if (position) positions.add(position);
  }

  const options = [...positions].sort((a, b) => a.localeCompare(b, "th"));
  if (options.length === 0) throw new Error("ไม่พบข้อมูลตำแหน่งใน Employee Master");
  return options;
}
