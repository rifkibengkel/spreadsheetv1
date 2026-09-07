'use client';

import React, { useState } from 'react';
import { useRouter } from 'next/navigation';

export default function LoginPage() {
  const router = useRouter();
  const [identifier, setIdentifier] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    setLoading(true);

    try {
      const res = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ identifier, password }),
        credentials: 'include',
      });

      const data = await res.json();

      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Login gagal.');
      }

      // Redirect to spreadsheet workspace
      router.push('/');
      router.refresh();
    } catch (err: any) {
      setError(err.message || 'Terjadi kesalahan saat login.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div
      style={{
        minHeight: '100vh',
        backgroundColor: '#090d16',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '24px',
        fontFamily: 'Inter, system-ui, sans-serif',
      }}
    >
      <div
        style={{
          width: '100%',
          maxWidth: '420px',
          backgroundColor: '#131b2e',
          border: '1px solid #232f48',
          borderRadius: '16px',
          padding: '40px 32px',
          boxShadow: '0 20px 40px rgba(0, 0, 0, 0.5)',
        }}
      >
        {/* Brand Header */}
        <div style={{ textAlign: 'center', marginBottom: '32px' }}>
          <div
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: '48px',
              height: '48px',
              borderRadius: '12px',
              background: 'linear-gradient(135deg, #0284c7 0%, #0369a1 100%)',
              color: '#ffffff',
              fontSize: '24px',
              marginBottom: '16px',
              boxShadow: '0 4px 12px rgba(2, 132, 199, 0.4)',
            }}
          >
            📊
          </div>
          <h1
            style={{
              color: '#f8fafc',
              fontSize: '22px',
              fontWeight: '700',
              margin: '0 0 6px 0',
              letterSpacing: '-0.02em',
            }}
          >
            Redbox Digital Sheet
          </h1>
          <p
            style={{
              color: '#94a3b8',
              fontSize: '13px',
              margin: 0,
            }}
          >
            Masuk untuk mengakses spreadsheet enterprise Anda
          </p>
        </div>

        {/* Error Alert */}
        {error && (
          <div
            style={{
              padding: '12px 14px',
              borderRadius: '8px',
              backgroundColor: 'rgba(153, 27, 27, 0.2)',
              border: '1px solid #991b1b',
              color: '#fca5a5',
              fontSize: '13px',
              marginBottom: '20px',
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
            }}
          >
            <span>⚠️</span>
            <span>{error}</span>
          </div>
        )}

        {/* Form */}
        <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '18px' }}>
          <div>
            <label
              htmlFor="identifier"
              style={{
                display: 'block',
                color: '#cbd5e1',
                fontSize: '13px',
                fontWeight: '600',
                marginBottom: '6px',
              }}
            >
              Username atau Email
            </label>
            <input
              id="identifier"
              type="text"
              required
              value={identifier}
              onChange={(e) => setIdentifier(e.target.value)}
              placeholder="Contoh: rootadmin atau user@perusahaan.com"
              style={{
                width: '100%',
                padding: '11px 14px',
                backgroundColor: '#0b1120',
                border: '1px solid #28354f',
                borderRadius: '8px',
                color: '#f8fafc',
                fontSize: '14px',
                outline: 'none',
                boxSizing: 'border-box',
                transition: 'border-color 0.15s ease',
              }}
            />
          </div>

          <div>
            <label
              htmlFor="password"
              style={{
                display: 'block',
                color: '#cbd5e1',
                fontSize: '13px',
                fontWeight: '600',
                marginBottom: '6px',
              }}
            >
              Password
            </label>
            <input
              id="password"
              type="password"
              required
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="••••••••••••"
              style={{
                width: '100%',
                padding: '11px 14px',
                backgroundColor: '#0b1120',
                border: '1px solid #28354f',
                borderRadius: '8px',
                color: '#f8fafc',
                fontSize: '14px',
                outline: 'none',
                boxSizing: 'border-box',
                transition: 'border-color 0.15s ease',
              }}
            />
          </div>

          <button
            type="submit"
            disabled={loading}
            style={{
              marginTop: '8px',
              padding: '12px 16px',
              backgroundColor: loading ? '#334155' : '#0284c7',
              color: '#ffffff',
              border: 'none',
              borderRadius: '8px',
              fontSize: '14px',
              fontWeight: '600',
              cursor: loading ? 'not-allowed' : 'pointer',
              transition: 'background-color 0.15s ease',
              boxShadow: '0 4px 12px rgba(2, 132, 199, 0.3)',
            }}
          >
            {loading ? 'Memverifikasi kredensial...' : 'Masuk ke Sistem'}
          </button>
        </form>

        {/* Footer info */}
        <div
          style={{
            marginTop: '28px',
            paddingTop: '20px',
            borderTop: '1px solid #1e293b',
            textAlign: 'center',
            fontSize: '12px',
            color: '#64748b',
          }}
        >
          Belum memiliki akun?{' '}
          <a
            href="/register"
            style={{
              color: '#38bdf8',
              textDecoration: 'none',
              fontWeight: '600',
            }}
          >
            Daftar Akun Baru
          </a>
        </div>
      </div>
    </div>
  );
}
