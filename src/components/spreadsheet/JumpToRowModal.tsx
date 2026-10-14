'use client';

import React, { useState, useEffect, useRef } from 'react';

interface JumpToRowModalProps {
  isOpen: boolean;
  onClose: () => void;
  maxRow: number;
  onJump: (row: number) => Promise<boolean | void>;
}

export default function JumpToRowModal({
  isOpen,
  onClose,
  maxRow,
  onJump,
}: JumpToRowModalProps) {
  const [targetRowStr, setTargetRowStr] = useState('1');
  const [isJumping, setIsJumping] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [jumpStatus, setJumpStatus] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isOpen) {
      setErrorMessage(null);
      setIsJumping(false);
      setJumpStatus(null);
      setTimeout(() => {
        inputRef.current?.focus();
        inputRef.current?.select();
      }, 60);
    }
  }, [isOpen]);

  if (!isOpen) return null;

  const parsedRow = parseInt(targetRowStr.replace(/[,._\s]/g, ''), 10);
  const isValidRow = !isNaN(parsedRow) && parsedRow >= 1 && parsedRow <= maxRow;

  // Calculate virtual window preview (1,000 rows window)
  const windowStart = isValidRow ? Math.max(1, parsedRow - 500) : 1;
  const windowEnd = isValidRow ? Math.min(maxRow, parsedRow + 500) : Math.min(maxRow, 1000);

  const handleExecuteJump = async (rowNum: number) => {
    if (isNaN(rowNum) || rowNum < 1 || rowNum > maxRow) {
      setErrorMessage(`Nomor baris harus antara 1 dan ${maxRow.toLocaleString()}.`);
      return;
    }

    setIsJumping(true);
    setErrorMessage(null);
    setJumpStatus(`⚡ Memuat virtual window baris ${rowNum.toLocaleString()}...`);

    const startTime = Date.now();
    try {
      await onJump(rowNum);
      
      // Ensure smooth visual perception of lazy load feedback
      const elapsed = Date.now() - startTime;
      if (elapsed < 350) {
        await new Promise((r) => setTimeout(r, 350 - elapsed));
      }
      onClose();
    } catch (err: any) {
      setErrorMessage(err?.message || 'Gagal melompat ke baris yang dituju.');
      setIsJumping(false);
      setJumpStatus(null);
    }
  };

  const onSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    handleExecuteJump(parsedRow);
  };

  const presets = [
    { label: 'Baris 1 (Awal)', row: 1 },
    { label: 'Baris 10.000', row: 10000 },
    { label: 'Baris 100.000', row: 100000 },
    { label: 'Baris 1.000.000', row: 1000000 },
    { label: 'Baris 4.000.000', row: 4000000 },
    { label: `Baris Terakhir (${maxRow.toLocaleString()})`, row: maxRow },
  ].filter((p) => p.row <= maxRow);

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 60,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        background: 'rgba(11, 15, 25, 0.78)',
        backdropFilter: 'blur(12px)',
        padding: '16px',
      }}
      onClick={(e) => {
        if (e.target === e.currentTarget && !isJumping) onClose();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        style={{
          width: '100%',
          maxWidth: '520px',
          background: '#0f172a',
          border: '1px solid #334155',
          borderRadius: '16px',
          boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.8), 0 0 0 1px rgba(255, 255, 255, 0.05)',
          overflow: 'hidden',
          display: 'flex',
          flexDirection: 'column',
          color: '#f8fafc',
          fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
        }}
      >
        {/* Header */}
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '18px 24px',
            borderBottom: '1px solid #1e293b',
            background: 'linear-gradient(180deg, rgba(30, 41, 59, 0.5) 0%, rgba(15, 23, 42, 0.5) 100%)',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div
              style={{
                width: '36px',
                height: '36px',
                borderRadius: '10px',
                background: 'linear-gradient(135deg, rgba(99, 102, 241, 0.25) 0%, rgba(56, 189, 248, 0.25) 100%)',
                border: '1px solid rgba(99, 102, 241, 0.4)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: '#818cf8',
                fontSize: '18px',
                boxShadow: '0 0 16px rgba(99, 102, 241, 0.2)',
              }}
            >
              🔍
            </div>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: '700', color: '#f8fafc' }}>
                  Lompat ke Baris
                </h3>
                <span
                  style={{
                    fontSize: '11px',
                    fontWeight: '600',
                    color: '#94a3b8',
                    background: '#1e293b',
                    border: '1px solid #334155',
                    padding: '1px 6px',
                    borderRadius: '4px',
                  }}
                >
                  Ctrl+G
                </span>
              </div>
              <p style={{ margin: '2px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                Navigasi instan hingga {maxRow.toLocaleString()} baris
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            disabled={isJumping}
            style={{
              background: 'transparent',
              border: 'none',
              color: '#64748b',
              fontSize: '18px',
              cursor: isJumping ? 'not-allowed' : 'pointer',
              padding: '6px',
              borderRadius: '8px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              transition: 'color 0.15s ease, background 0.15s ease',
            }}
            onMouseEnter={(e) => (e.currentTarget.style.color = '#f8fafc')}
            onMouseLeave={(e) => (e.currentTarget.style.color = '#64748b')}
          >
            ✕
          </button>
        </div>

        {/* Content Body */}
        <form onSubmit={onSubmit} style={{ padding: '22px 24px', display: 'flex', flexDirection: 'column', gap: '18px' }}>
          {/* Lazy Loading Architecture Callout */}
          <div
            style={{
              background: 'rgba(30, 41, 59, 0.45)',
              border: '1px solid rgba(56, 189, 248, 0.2)',
              borderRadius: '12px',
              padding: '12px 14px',
              display: 'flex',
              gap: '12px',
              alignItems: 'flex-start',
            }}
          >
            <span style={{ fontSize: '18px', lineHeight: 1 }}>⚡</span>
            <div style={{ flex: 1 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '4px' }}>
                <span style={{ fontSize: '12px', fontWeight: '700', color: '#38bdf8' }}>
                  Arsitektur Lazy Load & Virtual Windowing
                </span>
              </div>
              <p style={{ margin: 0, fontSize: '12px', color: '#94a3b8', lineHeight: 1.45 }}>
                Saat melompat, spreadsheet hanya memuat jendela virtual (~1.000 baris) di sekitar baris tujuan secara on-demand. Memori browser tetap sangat ringan (&lt; 2MB) dan bebas lag.
              </p>
            </div>
          </div>

          {/* Row Input Section */}
          <div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
              <label style={{ fontSize: '12px', fontWeight: '600', color: '#cbd5e1' }}>
                Masukkan Nomor Baris (1 – {maxRow.toLocaleString()})
              </label>
              {isValidRow && (
                <span style={{ fontSize: '11px', color: '#a5b4fc', fontWeight: '500' }}>
                  Window: {windowStart.toLocaleString()} – {windowEnd.toLocaleString()}
                </span>
              )}
            </div>

            <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
              <input
                ref={inputRef}
                type="text"
                value={targetRowStr}
                onChange={(e) => {
                  setTargetRowStr(e.target.value);
                  setErrorMessage(null);
                }}
                disabled={isJumping}
                placeholder="Contoh: 1000000"
                style={{
                  width: '100%',
                  height: '46px',
                  background: '#1e293b',
                  border: errorMessage ? '1px solid #f43f5e' : '1px solid #334155',
                  borderRadius: '10px',
                  padding: '0 16px',
                  fontSize: '16px',
                  color: '#f8fafc',
                  fontFamily: 'monospace',
                  fontWeight: '600',
                  outline: 'none',
                  boxShadow: 'inset 0 2px 4px rgba(0, 0, 0, 0.2)',
                  transition: 'border-color 0.15s ease, box-shadow 0.15s ease',
                }}
                onFocus={(e) => {
                  e.currentTarget.style.borderColor = '#6366f1';
                  e.currentTarget.style.boxShadow = '0 0 0 3px rgba(99, 102, 241, 0.2)';
                }}
                onBlur={(e) => {
                  e.currentTarget.style.borderColor = errorMessage ? '#f43f5e' : '#334155';
                  e.currentTarget.style.boxShadow = 'none';
                }}
              />
            </div>

            {errorMessage && (
              <div
                style={{
                  marginTop: '8px',
                  padding: '8px 12px',
                  borderRadius: '8px',
                  background: 'rgba(244, 63, 94, 0.12)',
                  border: '1px solid rgba(244, 63, 94, 0.3)',
                  color: '#fda4af',
                  fontSize: '12px',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                }}
              >
                <span>⚠️</span>
                <span>{errorMessage}</span>
              </div>
            )}
          </div>

          {/* Quick Presets */}
          <div>
            <span style={{ display: 'block', fontSize: '11px', fontWeight: '600', color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: '8px' }}>
              Pintasan Cepat:
            </span>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
              {presets.map((p) => (
                <button
                  key={p.row}
                  type="button"
                  disabled={isJumping}
                  onClick={() => {
                    setTargetRowStr(p.row.toString());
                    handleExecuteJump(p.row);
                  }}
                  style={{
                    fontSize: '12px',
                    fontWeight: '500',
                    background: '#1e293b',
                    color: '#cbd5e1',
                    border: '1px solid #334155',
                    borderRadius: '8px',
                    padding: '6px 12px',
                    cursor: isJumping ? 'not-allowed' : 'pointer',
                    transition: 'all 0.15s ease',
                  }}
                  onMouseEnter={(e) => {
                    if (!isJumping) {
                      e.currentTarget.style.background = '#334155';
                      e.currentTarget.style.color = '#f8fafc';
                      e.currentTarget.style.borderColor = '#475569';
                    }
                  }}
                  onMouseLeave={(e) => {
                    if (!isJumping) {
                      e.currentTarget.style.background = '#1e293b';
                      e.currentTarget.style.color = '#cbd5e1';
                      e.currentTarget.style.borderColor = '#334155';
                    }
                  }}
                >
                  {p.label}
                </button>
              ))}
            </div>
          </div>

          {/* Active Jumping Status / Progress */}
          {isJumping && (
            <div
              style={{
                background: 'rgba(30, 41, 59, 0.8)',
                border: '1px solid rgba(99, 102, 241, 0.3)',
                borderRadius: '10px',
                padding: '12px 14px',
                display: 'flex',
                flexDirection: 'column',
                gap: '8px',
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <span style={{ fontSize: '12px', fontWeight: '600', color: '#38bdf8' }}>
                  {jumpStatus || 'Melompat ke baris...'}
                </span>
                <span style={{ fontSize: '11px', color: '#94a3b8' }}>Streaming Window</span>
              </div>
              <div
                style={{
                  height: '4px',
                  width: '100%',
                  borderRadius: '2px',
                  overflow: 'hidden',
                  background: '#0f172a',
                }}
              >
                <div
                  className="animate-shimmer-lazy"
                  style={{
                    height: '100%',
                    width: '100%',
                    borderRadius: '2px',
                  }}
                />
              </div>
            </div>
          )}

          {/* Footer Actions */}
          <div
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'flex-end',
              gap: '10px',
              paddingTop: '16px',
              borderTop: '1px solid #1e293b',
            }}
          >
            <button
              type="button"
              onClick={onClose}
              disabled={isJumping}
              style={{
                background: 'transparent',
                color: '#94a3b8',
                border: '1px solid #334155',
                borderRadius: '8px',
                padding: '9px 18px',
                fontSize: '13px',
                fontWeight: '600',
                cursor: isJumping ? 'not-allowed' : 'pointer',
                transition: 'background 0.15s ease, color 0.15s ease',
              }}
              onMouseEnter={(e) => {
                if (!isJumping) {
                  e.currentTarget.style.background = '#1e293b';
                  e.currentTarget.style.color = '#f8fafc';
                }
              }}
              onMouseLeave={(e) => {
                if (!isJumping) {
                  e.currentTarget.style.background = 'transparent';
                  e.currentTarget.style.color = '#94a3b8';
                }
              }}
            >
              Batal
            </button>

            <button
              type="submit"
              disabled={isJumping || !isValidRow}
              style={{
                background: isJumping || !isValidRow ? '#334155' : 'linear-gradient(135deg, #4f46e5 0%, #6366f1 100%)',
                color: isJumping || !isValidRow ? '#64748b' : '#ffffff',
                border: 'none',
                borderRadius: '8px',
                padding: '9px 22px',
                fontSize: '13px',
                fontWeight: '600',
                cursor: isJumping || !isValidRow ? 'not-allowed' : 'pointer',
                boxShadow: isJumping || !isValidRow ? 'none' : '0 4px 14px rgba(79, 70, 229, 0.4)',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                transition: 'all 0.15s ease',
              }}
            >
              {isJumping ? (
                <>
                  <span style={{ display: 'inline-block', animation: 'spin-slow 1s infinite linear' }}>⚡</span>
                  <span>Melompat (Lazy Load)...</span>
                </>
              ) : (
                <>
                  <span>Lompat ke Baris</span>
                  <span>→</span>
                </>
              )}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
