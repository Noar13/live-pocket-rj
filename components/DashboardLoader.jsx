'use client';
import dynamic from 'next/dynamic';

// Dashboard murni client-side (Firebase Auth, tanggal, interval pasar) -> render hanya di browser
const CashflowDashboard = dynamic(() => import('./CashflowDashboard'), {
  ssr: false,
  loading: () => (
    <div className="flex min-h-screen items-center justify-center text-sm opacity-70">
      Memuat DompetKu...
    </div>
  ),
});

export default function DashboardLoader() {
  return <CashflowDashboard />;
}
