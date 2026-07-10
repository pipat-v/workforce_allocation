alter table public.leave_records
drop constraint if exists leave_records_leave_type_check;

alter table public.leave_records
add constraint leave_records_leave_type_check
check (leave_type in (
  'ลาป่วย',
  'ลากิจ',
  'ลาพักร้อน',
  'ขาดงาน',
  'ลาตรวจครรภ์',
  'ลาคลอด',
  'ลาคลอดคู่สมรส',
  'ลาอุบัติเหตุจากการปฏิบัติงาน',
  'ลาบวช/ลาพิธีสำคัญทางศาสนา',
  'ลาทหาร',
  'ลาพิเศษไม่จ่าย',
  'สแกนหน้าไม่ติด'
));
