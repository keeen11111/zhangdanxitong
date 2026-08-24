import Link from "next/link";
import { ArrowLeft } from "lucide-react";

export default function ProjectLayout({ children }: { children: React.ReactNode }) {
  return (
    <div>
      <div className="mx-auto mb-5 max-w-5xl">
        <Link
          href="/projects"
          className="inline-flex items-center gap-1.5 text-sm text-slate-500 hover:text-slate-900 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-sky-700"
        >
          <ArrowLeft className="h-4 w-4" />返回整合任务
        </Link>
      </div>
      {children}
    </div>
  );
}
