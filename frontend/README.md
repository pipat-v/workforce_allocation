# Workforce Allocation Frontend

Next.js frontend for Vercel with Supabase Auth, Database, and Storage.

## Local setup

1. Create a Supabase project.
2. Run `supabase/schema.sql` in the Supabase SQL editor.
3. Copy `.env.example` to `.env.local`.
4. Fill:
   - `NEXT_PUBLIC_SUPABASE_URL`
   - `NEXT_PUBLIC_SUPABASE_ANON_KEY`
5. Install and run:

```bash
npm install
npm run dev
```

The frontend uploads input files to the private `workforce-inputs` bucket and creates a row in `allocation_runs`.

## Vercel setup

Set the same environment variables in Vercel Project Settings, then deploy the `frontend` directory.

## Processing worker

The Python allocation logic should run as a separate worker/API service. It should:

1. Poll or receive a request for `allocation_runs.status = 'uploaded'`.
2. Download the four input files from Supabase Storage.
3. Call the existing Python pipeline logic.
4. Upload the result Excel file to `workforce-outputs`.
5. Insert rows into `allocation_results` and `gap_summaries`.
6. Update `allocation_runs.status` to `completed` or `failed`.

