import { AuthGuard } from "@/components/auth-guard";
import { SaaSShell } from "@/components/saas-shell";

export default function DashboardLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <AuthGuard>
      <SaaSShell>{children}</SaaSShell>
    </AuthGuard>
  );
}
