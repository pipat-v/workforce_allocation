# Supabase + Vercel Architecture

## Target Architecture

- Vercel hosts the Next.js frontend.
- Supabase handles Auth, Postgres, Row Level Security, and file storage.
- Python runs the allocation solver as a separate worker/API service.

The Python solver uses `pandas`, `openpyxl`, and `pulp`, so it should not be embedded directly in browser code. Vercel Functions can run Python, but allocation jobs can exceed request/function limits as data grows. A worker service is safer for production.

## Data Flow

1. User logs in on the frontend.
2. User uploads:
   - Timestamp / Time Record
   - Master Employee
   - Manpower Plan
   - Skill Matrix
3. Files are stored in Supabase Storage under:
   - `workforce-inputs/{user_id}/{run_id}/...`
4. Frontend creates one row in `allocation_runs`.
5. Worker picks up the run.
6. Worker downloads files, runs allocation, and writes:
   - Excel output to `workforce-outputs`
   - detailed assignment rows to `allocation_results`
   - shortage/surplus rows to `gap_summaries`
7. Frontend displays status and results.

## Implementation Phases

1. Create Supabase project and run `supabase/schema.sql`.
2. Deploy `frontend/` to Vercel.
3. Add Supabase Auth UI or invite-only login.
4. Build worker integration around the existing Python pipeline.
5. Add result detail pages and Excel download.
6. Add admin views for all runs and user management.

