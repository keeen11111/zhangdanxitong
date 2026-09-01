export default function DashboardLoading() {
  return (
    <div className="payroll-page animate-pulse" aria-busy="true" aria-label="正在加载页面">
      <header className="border-b border-slate-200 pb-5">
        <div className="h-3 w-24 rounded bg-slate-200" />
        <div className="mt-3 h-8 w-44 rounded bg-slate-200" />
        <div className="mt-3 h-4 w-72 max-w-full rounded bg-slate-100" />
      </header>
      <section className="rounded-md border border-slate-200 bg-white p-5">
        <div className="h-5 w-36 rounded bg-slate-200" />
        <div className="mt-5 grid gap-3 sm:grid-cols-3">
          {[0, 1, 2].map((item) => <div key={item} className="h-20 rounded-md bg-slate-100" />)}
        </div>
      </section>
      <section className="rounded-md border border-slate-200 bg-white p-5">
        <div className="h-5 w-28 rounded bg-slate-200" />
        <div className="mt-5 space-y-3">
          {[0, 1, 2, 3].map((item) => <div key={item} className="h-12 rounded-md bg-slate-100" />)}
        </div>
      </section>
    </div>
  );
}
