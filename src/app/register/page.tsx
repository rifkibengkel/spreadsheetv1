'use client';

import React, { useState } from 'react';
import Link from 'next/link';

export default function RegisterPage() {
  const [username, setUsername] = useState('');
  const [email, setEmail] = useState('');
  const [fullName, setFullName] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);

    if (password !== confirmPassword) {
      setError('Konfirmasi password tidak cocok.');
      return;
    }

    if (password.length < 8) {
      setError('Password minimal 8 karakter.');
      return;
    }

    setLoading(true);

    try {
      const res = await fetch('/api/auth/register', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, email, password, fullName }),
      });

      const data = await res.json();

      if (!res.ok || !data.success) {
        throw new Error(data.error || 'Pendaftaran gagal.');
      }

      setSuccess(true);
    } catch (err: any) {
      setError(err.message || 'Terjadi kesalahan saat mendaftar.');
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
          maxWidth: '460px',
          backgroundColor: '#131b2e',
          border: '1px solid #232f48',
          borderRadius: '16px',
          padding: '40px 32px',
          boxShadow: '0 20px 40px rgba(0, 0, 0, 0.5)',
        }}
      >
        {/* Brand Header */}
        <div style={{ textAlign: 'center', marginBottom: '28px' }}>
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
            }}
          >
            📋
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
            Pendaftaran Pengguna Baru
          </h1>
          <p
            style={{
              color: '#94a3b8',
              fontSize: '13px',
              margin: 0,
            }}
          >
            Akun Anda akan ditinjau oleh Administrator sebelum aktif
          </p>
        </div>

        {success ? (
          <div
            style={{
              padding: '24px',
              borderRadius: '12px',
              backgroundColor: 'rgba(6, 78, 59, 0.3)',
              border: '1px solid #059669',
              textAlign: 'center',
            }}
          >
            <div style={{ fontSize: '36px', marginBottom: '12px' }}>✅</div>
            <h3 style={{ color: '#6ee7b7', margin: '0 0 8px 0', fontSize: '16px', fontWeight: '700' }}>
              Pendaftaran Berhasil!
            </h3>
            <p style={{ color: '#a7f3d0', fontSize: '13px', lineHeight: '1.5', margin: '0 0 20px 0' }}>
              Permohonan akun Anda telah dicatat dengan status <strong>PENDING</strong>. Anda dapat masuk setelah disetujui oleh Administrator.
            </p>
            <Link
              href="/login"
              style={{
                display: 'inline-block',
                padding: '10px 20px',
                backgroundColor: '#059669',
                color: '#ffffff',
                borderRadius: '8px',
                textDecoration: 'none',
                fontWeight: '600',
                fontSize: '13px',
              }}
            >
              Kembali ke Halaman Login
            </Link>
          </div>
        ) : (
          <>
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
                }}
              >
                ⚠️ {error}
              </div>
            )}

            <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
              <div>
                <label
                  htmlFor="fullName"
                  style={{ display: 'block', color: '#cbd5e1', fontSize: '13px', fontWeight: '600', marginBottom: '5px' }}
                >
                  Nama Lengkap
                </label>
                <input
                  id="fullName"
                  type="text"
                  required
                  value={fullName}
                  onChange={(e) => setFullName(e.target.value)}
                  placeholder="Contoh: Budi Pratama"
                  style={inputStyle}
                />
              </div>

              <div>
                <label
                  htmlFor="username"
                  style={{ display: 'block', color: '#cbd5e1', fontSize: '13px', fontWeight: '600', marginBottom: '5px' }}
                >
                  Username
                </label>
                <input
                  id="username"
                  type="text"
                  required
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="Contoh: budipratama"
                  style={inputStyle}
                />
              </div>

              <div>
                <label
                  htmlFor="email"
                  style={{ display: 'block', color: '#cbd5e1', fontSize: '13px', fontWeight: '600', marginBottom: '5px' }}
                >
                  Email Perusahaan
                </label>
                <input
                  id="email"
                  type="email"
                  required
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="budi@perusahaan.com"
                  style={inputStyle}
                />
              </div>

              <div>
                <label
                  htmlFor="password"
                  style={{ display: 'block', color: '#cbd5e1', fontSize: '13px', fontWeight: '600', marginBottom: '5px' }}
                >
                  Password (Minimal 8 Karakter)
                </label>
                <input
                  id="password"
                  type="password"
                  required
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••••••"
                  style={inputStyle}
                />
              </div>

              <div>
                <label
                  htmlFor="confirmPassword"
                  style={{ display: 'block', color: '#cbd5e1', fontSize: '13px', fontWeight: '600', marginBottom: '5px' }}
                >
                  Konfirmasi Password
                </label>
                <input
                  id="confirmPassword"
                  type="password"
                  required
                  value={confirmPassword}
                  onChange={(e) => setConfirmPassword(e.target.value)}
                  placeholder="••••••••••••"
                  style={inputStyle}
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
                  boxShadow: '0 4px 12px rgba(2, 132, 199, 0.3)',
                }}
              >
                {loading ? 'Mendaftarkan Akun...' : 'Kirim Pendaftaran'}
              </button>
            </form>

            <div
              style={{
                marginTop: '24px',
                paddingTop: '18px',
                borderTop: '1px solid #1e293b',
                textAlign: 'center',
                fontSize: '12px',
                color: '#64748b',
              }}
            >
              Sudah memiliki akun terdaftar?{' '}
              <Link href="/login" style={{ color: '#38bdf8', textDecoration: 'none', fontWeight: '600' }}>
                Masuk di sini
              </Link>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

const inputStyle: React.CSSProperties = {
  width: '100%',
  padding: '10px 14px',
  backgroundColor: '#0b1120',
  border: '1px solid #28354f',
  borderRadius: '8px',
  color: '#f8fafc',
  fontSize: '14px',
  outline: 'none',
  boxSizing: 'border-box',
};
