create extension if not exists pg_cron with schema pg_catalog;

do $$
begin
  if exists (select 1 from cron.job where jobname = 'refresh-production-user-daily') then
    perform cron.unschedule('refresh-production-user-daily');
  end if;
end
$$;

select cron.schedule(
  'refresh-production-user-daily',
  '5 17 * * *',
  $cron$select public.refresh_production_user((now() at time zone 'Asia/Bangkok')::date);$cron$
);
