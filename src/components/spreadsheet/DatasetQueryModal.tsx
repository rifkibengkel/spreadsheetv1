'use client';

import React, { useState, useEffect } from 'react';
import {
  QueryDefinition,
  FilterPredicate,
  FilterOperator,
  SortDirection,
} from '@/lib/largeCsv/query/types';

interface DatasetQueryModalProps {
  isOpen: boolean;
  onClose: () => void;
  headers: string[];
  datasetId: string;
  datasetRowCount?: number;
  activeQueryDef?: QueryDefinition | null;
  onApplyQuery: (queryDef: QueryDefinition) => Promise<void>;
  isLoading?: boolean;
}

const OPERATOR_OPTIONS: { label: string; value: FilterOperator }[] = [
  { label: 'Sama Dengan (=)', value: 'EQUALS' },
  { label: 'Tidak Sama (!=)', value: 'NOT_EQUALS' },
  { label: 'Mengandung Teks', value: 'CONTAINS' },
  { label: 'Diawali Dengan', value: 'STARTS_WITH' },
  { label: 'Diakhiri Dengan', value: 'ENDS_WITH' },
  { label: 'Lebih Besar (>)', value: 'GREATER_THAN' },
  { label: 'Lebih Kecil (<)', value: 'LESS_THAN' },
  { label: 'Kosong (Empty)', value: 'IS_EMPTY' },
  { label: 'Tidak Kosong (Not Empty)', value: 'IS_NOT_EMPTY' },
  { label: 'Di Antara Pilihan (IN)', value: 'IN' },
];

export default function DatasetQueryModal({
  isOpen,
  onClose,
  headers,
  datasetId,
  datasetRowCount,
  activeQueryDef,
  onApplyQuery,
  isLoading = false,
}: DatasetQueryModalProps) {
  const [conjunction, setConjunction] = useState<'AND' | 'OR'>('AND');
  const [predicates, setPredicates] = useState<FilterPredicate[]>([]);
  const [hasSort, setHasSort] = useState(false);
  const [sortColumn, setSortColumn] = useState('');
  const [sortDirection, setSortDirection] = useState<SortDirection>('ASC');
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [elapsedSeconds, setElapsedSeconds] = useState(0);

  // Real-time elapsed timer while executing lazy query
  useEffect(() => {
    let timer: any;
    if (isLoading) {
      setElapsedSeconds(0);
      timer = setInterval(() => {
        setElapsedSeconds((s) => s + 1);
      }, 1000);
    } else {
      setElapsedSeconds(0);
    }
    return () => clearInterval(timer);
  }, [isLoading]);

  useEffect(() => {
    if (isOpen) {
      setErrorMessage(null);
      if (activeQueryDef) {
        setConjunction(activeQueryDef.filter?.conjunction || 'AND');
        setPredicates(activeQueryDef.filter?.predicates || []);
        if (activeQueryDef.sort?.columns && activeQueryDef.sort.columns.length > 0) {
          setHasSort(true);
          setSortColumn(activeQueryDef.sort.columns[0].column);
          setSortDirection(activeQueryDef.sort.columns[0].direction);
        } else {
          setHasSort(false);
          setSortColumn(headers[0] || '');
          setSortDirection('ASC');
        }
      } else {
        setConjunction('AND');
        setPredicates([]);
        setHasSort(false);
        setSortColumn(headers[0] || '');
        setSortDirection('ASC');
      }
    }
  }, [isOpen, activeQueryDef, headers]);

  if (!isOpen) return null;

  const handleAddPredicate = () => {
    const defaultCol = headers.find((h) => h === 'Registration_State') || headers[0] || '';
    setPredicates((prev) => [
      ...prev,
      {
        column: defaultCol,
        operator: 'EQUALS',
        value: '',
        caseSensitive: false,
        numeric: false,
      },
    ]);
  };

  const handleRemovePredicate = (index: number) => {
    setPredicates((prev) => prev.filter((_, i) => i !== index));
  };

  const handleUpdatePredicate = (index: number, patch: Partial<FilterPredicate>) => {
    setPredicates((prev) =>
      prev.map((p, i) => (i === index ? { ...p, ...patch } : p))
    );
  };

  const handleApplyPreset = (presetType: string) => {
    if (presetType === 'CT') {
      setPredicates([
        { column: 'Registration_State', operator: 'EQUALS', value: 'CT' },
      ]);
      setHasSort(false);
    } else if (presetType === 'NY') {
      setPredicates([
        { column: 'Registration_State', operator: 'EQUALS', value: 'NY' },
      ]);
      setHasSort(false);
    } else if (presetType === 'SUMMONS') {
      setPredicates([
        { column: 'Summons_Number', operator: 'EQUALS', value: '8002531292' },
      ]);
      setHasSort(false);
    } else if (presetType === 'SORT_DATE') {
      setPredicates([
        { column: 'Registration_State', operator: 'EQUALS', value: 'FL' },
      ]);
      setHasSort(true);
      setSortColumn('Issue_Date');
      setSortDirection('ASC');
    } else if (headers.length > 0) {
      // Generic preset using first header
      setPredicates([
        { column: headers[0], operator: 'IS_NOT_EMPTY', value: '' },
      ]);
      setHasSort(true);
      setSortColumn(headers[0]);
      setSortDirection('ASC');
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setErrorMessage(null);

    const activePreds = predicates.filter((p) => {
      if (p.operator === 'IS_EMPTY' || p.operator === 'IS_NOT_EMPTY') return true;
      return p.value !== undefined && p.value !== '';
    });

    if (activePreds.length === 0 && !hasSort) {
      setErrorMessage('Harap tentukan setidaknya satu kriteria filter atau aktifkan pengurutan kolom.');
      return;
    }

    const queryDef: QueryDefinition = {};

    if (activePreds.length > 0) {
      queryDef.filter = {
        conjunction,
        predicates: activePreds,
      };
    }

    if (hasSort && sortColumn) {
      queryDef.sort = {
        columns: [{ column: sortColumn, direction: sortDirection }],
      };
    }

    try {
      await onApplyQuery(queryDef);
      onClose();
    } catch (err: any) {
      setErrorMessage(err?.message || 'Gagal menjalankan query server. Pastikan koneksi atau parameter valid.');
    }
  };

  const hasParkingHeaders = headers.includes('Registration_State') || headers.includes('Summons_Number');

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 60,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        background: 'rgba(11, 15, 25, 0.82)',
        backdropFilter: 'blur(14px)',
        padding: '16px',
      }}
      onClick={(e) => {
        if (e.target === e.currentTarget && !isLoading) onClose();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        style={{
          width: '100%',
          maxWidth: '720px',
          maxHeight: '90vh',
          background: '#0f172a',
          border: '1px solid #334155',
          borderRadius: '16px',
          boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.85), 0 0 0 1px rgba(255, 255, 255, 0.05)',
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
            background: 'linear-gradient(180deg, rgba(30, 41, 59, 0.6) 0%, rgba(15, 23, 42, 0.6) 100%)',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div
              style={{
                width: '38px',
                height: '38px',
                borderRadius: '10px',
                background: 'linear-gradient(135deg, rgba(56, 189, 248, 0.25) 0%, rgba(99, 102, 241, 0.25) 100%)',
                border: '1px solid rgba(56, 189, 248, 0.4)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: '#38bdf8',
                fontSize: '18px',
                boxShadow: '0 0 16px rgba(56, 189, 248, 0.2)',
              }}
            >
              ⚡
            </div>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: '700', color: '#f8fafc' }}>
                  Filter & Sort Server
                </h3>
                <span
                  style={{
                    fontSize: '11px',
                    fontWeight: '700',
                    color: '#38bdf8',
                    background: 'rgba(56, 189, 248, 0.15)',
                    border: '1px solid rgba(56, 189, 248, 0.3)',
                    padding: '2px 8px',
                    borderRadius: '12px',
                  }}
                >
                  ⚡ Lazy Stream
                </span>
                {datasetRowCount && datasetRowCount > 0 ? (
                  <span
                    style={{
                      fontSize: '11px',
                      fontWeight: '600',
                      color: '#a5b4fc',
                      background: 'rgba(99, 102, 241, 0.15)',
                      border: '1px solid rgba(99, 102, 241, 0.3)',
                      padding: '2px 8px',
                      borderRadius: '12px',
                    }}
                  >
                    {datasetRowCount.toLocaleString()} baris
                  </span>
                ) : null}
              </div>
              <p style={{ margin: '2px 0 0', fontSize: '12px', color: '#94a3b8' }}>
                Menyaring jutaan baris server-side tanpa membebani RAM browser
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            disabled={isLoading}
            style={{
              background: 'transparent',
              border: 'none',
              color: '#64748b',
              fontSize: '18px',
              cursor: isLoading ? 'not-allowed' : 'pointer',
              padding: '6px',
              borderRadius: '8px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              transition: 'color 0.15s ease',
            }}
            onMouseEnter={(e) => (e.currentTarget.style.color = '#f8fafc')}
            onMouseLeave={(e) => (e.currentTarget.style.color = '#64748b')}
          >
            ✕
          </button>
        </div>

        {/* Form Body */}
        <form
          onSubmit={handleSubmit}
          style={{
            padding: '22px 24px',
            overflowY: 'auto',
            display: 'flex',
            flexDirection: 'column',
            gap: '20px',
          }}
        >
          {/* Lazy Loading Architecture Callout */}
          <div
            style={{
              background: 'rgba(15, 23, 42, 0.6)',
              border: '1px solid rgba(56, 189, 248, 0.25)',
              borderRadius: '12px',
              padding: '14px 16px',
              display: 'flex',
              gap: '12px',
              alignItems: 'flex-start',
            }}
          >
            <span style={{ fontSize: '20px', lineHeight: 1 }}>⚡</span>
            <div style={{ flex: 1 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px' }}>
                <span style={{ fontSize: '12px', fontWeight: '700', color: '#38bdf8' }}>
                  Arsitektur Lazy Load & Server Streaming
                </span>
                <span style={{ fontSize: '10px', color: '#10b981', background: 'rgba(16, 185, 129, 0.15)', border: '1px solid rgba(16, 185, 129, 0.3)', padding: '1px 6px', borderRadius: '4px', fontWeight: '600' }}>
                  RAM &lt; 5MB
                </span>
              </div>
              <p style={{ margin: 0, fontSize: '12px', color: '#94a3b8', lineHeight: 1.5 }}>
                Filter dan pengurutan dieksekusi langsung pada engine berkinerja tinggi. Hasil jutaan baris tidak di-download sekaligus, melainkan dialirkan secara bertahap (*lazy windowing 1.000 baris*) saat Anda melihat atau menggulir spreadsheet.
              </p>
            </div>
          </div>

          {/* Quick Presets */}
          <div
            style={{
              background: 'rgba(30, 41, 59, 0.5)',
              border: '1px solid #334155',
              borderRadius: '12px',
              padding: '12px 14px',
            }}
          >
            <span style={{ display: 'block', fontSize: '11px', fontWeight: '700', color: '#cbd5e1', textTransform: 'uppercase', letterSpacing: '0.05em', marginBottom: '8px' }}>
              ⚡ Pintasan Filter Cepat:
            </span>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
              {hasParkingHeaders ? (
                <>
                  <button
                    type="button"
                    disabled={isLoading}
                    onClick={() => handleApplyPreset('CT')}
                    style={presetBtnStyle}
                  >
                    Plat CT (~160k Baris)
                  </button>
                  <button
                    type="button"
                    disabled={isLoading}
                    onClick={() => handleApplyPreset('NY')}
                    style={presetBtnStyle}
                  >
                    Plat NY (Mayoritas)
                  </button>
                  <button
                    type="button"
                    disabled={isLoading}
                    onClick={() => handleApplyPreset('SUMMONS')}
                    style={presetBtnStyle}
                  >
                    Summons 8002531292 (1 Baris)
                  </button>
                  <button
                    type="button"
                    disabled={isLoading}
                    onClick={() => handleApplyPreset('SORT_DATE')}
                    style={presetBtnStyle}
                  >
                    Plat FL + Urut Tanggal ASC
                  </button>
                </>
              ) : (
                headers.slice(0, 3).map((h, i) => (
                  <button
                    key={h}
                    type="button"
                    disabled={isLoading}
                    onClick={() => {
                      setPredicates([
                        { column: h, operator: 'IS_NOT_EMPTY', value: '' },
                      ]);
                      setHasSort(true);
                      setSortColumn(h);
                      setSortDirection('ASC');
                    }}
                    style={presetBtnStyle}
                  >
                    Filter {h || `Kolom ${i + 1}`} (Tidak Kosong)
                  </button>
                ))
              )}
            </div>
          </div>

          {/* Filter Predicates Section */}
          <div>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '10px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <span style={{ fontSize: '13px', fontWeight: '700', color: '#f8fafc' }}>
                  Kriteria Filter
                </span>
                {predicates.length > 1 && (
                  <div style={{ display: 'flex', alignItems: 'center', background: '#1e293b', border: '1px solid #334155', borderRadius: '6px', padding: '2px' }}>
                    <button
                      type="button"
                      disabled={isLoading}
                      onClick={() => setConjunction('AND')}
                      style={{
                        padding: '3px 8px',
                        borderRadius: '4px',
                        border: 'none',
                        fontSize: '11px',
                        fontWeight: '600',
                        cursor: 'pointer',
                        background: conjunction === 'AND' ? '#4f46e5' : 'transparent',
                        color: conjunction === 'AND' ? '#ffffff' : '#94a3b8',
                      }}
                    >
                      AND (Semua)
                    </button>
                    <button
                      type="button"
                      disabled={isLoading}
                      onClick={() => setConjunction('OR')}
                      style={{
                        padding: '3px 8px',
                        borderRadius: '4px',
                        border: 'none',
                        fontSize: '11px',
                        fontWeight: '600',
                        cursor: 'pointer',
                        background: conjunction === 'OR' ? '#4f46e5' : 'transparent',
                        color: conjunction === 'OR' ? '#ffffff' : '#94a3b8',
                      }}
                    >
                      OR (Salah Satu)
                    </button>
                  </div>
                )}
              </div>

              <button
                type="button"
                disabled={isLoading}
                onClick={handleAddPredicate}
                style={{
                  background: 'rgba(99, 102, 241, 0.2)',
                  color: '#a5b4fc',
                  border: '1px solid rgba(99, 102, 241, 0.4)',
                  borderRadius: '6px',
                  padding: '5px 12px',
                  fontSize: '12px',
                  fontWeight: '600',
                  cursor: isLoading ? 'not-allowed' : 'pointer',
                  transition: 'all 0.15s ease',
                }}
              >
                + Tambah Filter
              </button>
            </div>

            {predicates.length === 0 ? (
              <div
                style={{
                  textAlign: 'center',
                  padding: '24px 16px',
                  border: '1px dashed #334155',
                  borderRadius: '10px',
                  color: '#64748b',
                  fontSize: '13px',
                }}
              >
                Belum ada filter yang ditambahkan. Klik <strong>&quot;+ Tambah Filter&quot;</strong> atau pilih <strong>Pintasan Filter Cepat</strong> di atas.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                {predicates.map((pred, idx) => (
                  <div
                    key={idx}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: '8px',
                      background: '#1e293b',
                      border: '1px solid #334155',
                      borderRadius: '10px',
                      padding: '10px 12px',
                      flexWrap: 'wrap',
                    }}
                  >
                    {/* Column Select */}
                    <select
                      value={pred.column}
                      disabled={isLoading}
                      onChange={(e) => handleUpdatePredicate(idx, { column: e.target.value })}
                      style={selectFieldStyle}
                    >
                      {headers.map((h, hIdx) => (
                        <option key={h || hIdx} value={h}>
                          {h ? `${hIdx + 1}. ${h}` : `(Kolom ${hIdx + 1})`}
                        </option>
                      ))}
                    </select>

                    {/* Operator Select */}
                    <select
                      value={pred.operator}
                      disabled={isLoading}
                      onChange={(e) =>
                        handleUpdatePredicate(idx, {
                          operator: e.target.value as FilterOperator,
                        })
                      }
                      style={selectFieldStyle}
                    >
                      {OPERATOR_OPTIONS.map((op) => (
                        <option key={op.value} value={op.value}>
                          {op.label}
                        </option>
                      ))}
                    </select>

                    {/* Value Input */}
                    {pred.operator !== 'IS_EMPTY' && pred.operator !== 'IS_NOT_EMPTY' && (
                      <input
                        type="text"
                        value={Array.isArray(pred.value) ? pred.value.join(', ') : pred.value || ''}
                        disabled={isLoading}
                        onChange={(e) => handleUpdatePredicate(idx, { value: e.target.value })}
                        placeholder="Nilai yang dicari..."
                        style={inputFieldStyle}
                      />
                    )}

                    {/* Delete button */}
                    <button
                      type="button"
                      disabled={isLoading}
                      onClick={() => handleRemovePredicate(idx)}
                      style={{
                        background: 'transparent',
                        border: 'none',
                        color: '#94a3b8',
                        cursor: isLoading ? 'not-allowed' : 'pointer',
                        padding: '6px',
                        borderRadius: '6px',
                        fontSize: '14px',
                        lineHeight: 1,
                        transition: 'color 0.15s ease',
                      }}
                      onMouseEnter={(e) => (e.currentTarget.style.color = '#f43f5e')}
                      onMouseLeave={(e) => (e.currentTarget.style.color = '#94a3b8')}
                      title="Hapus kriteria ini"
                    >
                      🗑️
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Sort Section */}
          <div style={{ borderTop: '1px solid #1e293b', paddingTop: '16px' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', marginBottom: '10px' }}>
              <input
                type="checkbox"
                checked={hasSort}
                disabled={isLoading}
                onChange={(e) => setHasSort(e.target.checked)}
                style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#4f46e5' }}
              />
              <span style={{ fontSize: '13px', fontWeight: '700', color: '#f8fafc' }}>
                Aktifkan Pengurutan Server (Sort)
              </span>
            </label>

            {hasSort && (
              <div
                style={{
                  display: 'flex',
                  gap: '12px',
                  background: '#1e293b',
                  border: '1px solid #334155',
                  borderRadius: '10px',
                  padding: '12px 14px',
                }}
              >
                <div style={{ flex: 1 }}>
                  <label style={{ display: 'block', fontSize: '11px', fontWeight: '600', color: '#94a3b8', marginBottom: '4px' }}>
                    Kolom Pengurutan:
                  </label>
                  <select
                    value={sortColumn}
                    disabled={isLoading}
                    onChange={(e) => setSortColumn(e.target.value)}
                    style={{ ...selectFieldStyle, width: '100%', maxWidth: 'none' }}
                  >
                    {headers.map((h, hIdx) => (
                      <option key={h || hIdx} value={h}>
                        {h ? `${hIdx + 1}. ${h}` : `(Kolom ${hIdx + 1})`}
                      </option>
                    ))}
                  </select>
                </div>

                <div style={{ width: '200px' }}>
                  <label style={{ display: 'block', fontSize: '11px', fontWeight: '600', color: '#94a3b8', marginBottom: '4px' }}>
                    Arah Urutan:
                  </label>
                  <select
                    value={sortDirection}
                    disabled={isLoading}
                    onChange={(e) => setSortDirection(e.target.value as SortDirection)}
                    style={{ ...selectFieldStyle, width: '100%', maxWidth: 'none' }}
                  >
                    <option value="ASC">Menaik (ASC / A → Z)</option>
                    <option value="DESC">Menurun (DESC / Z → A)</option>
                  </select>
                </div>
              </div>
            )}
          </div>

          {/* Error Message */}
          {errorMessage && (
            <div
              style={{
                padding: '10px 14px',
                borderRadius: '8px',
                background: 'rgba(244, 63, 94, 0.12)',
                border: '1px solid rgba(244, 63, 94, 0.3)',
                color: '#fda4af',
                fontSize: '12px',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
              }}
            >
              <span>⚠️</span>
              <span>{errorMessage}</span>
            </div>
          )}

          {/* Live Lazy Loading Progress Feedback Banner */}
          {isLoading && (
            <div
              style={{
                background: 'rgba(30, 41, 59, 0.9)',
                border: '1px solid rgba(56, 189, 248, 0.4)',
                borderRadius: '12px',
                padding: '16px',
                display: 'flex',
                flexDirection: 'column',
                gap: '10px',
                boxShadow: '0 8px 24px rgba(0, 0, 0, 0.4)',
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{ fontSize: '16px', animation: 'spin-slow 1.5s infinite linear' }}>⚡</span>
                  <span style={{ fontSize: '13px', fontWeight: '700', color: '#38bdf8' }}>
                    Memproses Query Server & Menyiapkan Lazy Stream...
                  </span>
                </div>
                <span style={{ fontSize: '12px', fontWeight: '700', color: '#fbbf24', fontFamily: 'monospace' }}>
                  ⏱️ {elapsedSeconds}s berjalan
                </span>
              </div>

              {/* Animated multi-gradient lazy bar */}
              <div
                style={{
                  height: '6px',
                  width: '100%',
                  borderRadius: '3px',
                  overflow: 'hidden',
                  background: '#0f172a',
                }}
              >
                <div
                  className="animate-shimmer-lazy"
                  style={{
                    height: '100%',
                    width: '100%',
                    borderRadius: '3px',
                  }}
                />
              </div>

              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '11px', color: '#94a3b8' }}>
                <span>🔍 Memindai jutaan baris server-side</span>
                <span>⚡ Menyiapkan jendela virtual 1.000 baris</span>
                <span>📊 Sinkronisasi Univer Canvas</span>
              </div>
              <p style={{ margin: 0, fontSize: '11px', color: '#64748b', fontStyle: 'italic', textAlign: 'center' }}>
                Harap tunggu, jutaan data disaring dengan presisi tanpa membuat browser freeze.
              </p>
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
              disabled={isLoading}
              style={{
                background: 'transparent',
                color: '#94a3b8',
                border: '1px solid #334155',
                borderRadius: '8px',
                padding: '9px 18px',
                fontSize: '13px',
                fontWeight: '600',
                cursor: isLoading ? 'not-allowed' : 'pointer',
                transition: 'all 0.15s ease',
              }}
            >
              Batal
            </button>

            <button
              type="submit"
              disabled={isLoading}
              style={{
                background: isLoading ? '#334155' : 'linear-gradient(135deg, #0284c7 0%, #4f46e5 100%)',
                color: isLoading ? '#94a3b8' : '#ffffff',
                border: 'none',
                borderRadius: '8px',
                padding: '9px 24px',
                fontSize: '13px',
                fontWeight: '600',
                cursor: isLoading ? 'not-allowed' : 'pointer',
                boxShadow: isLoading ? 'none' : '0 4px 14px rgba(2, 132, 199, 0.4)',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                transition: 'all 0.15s ease',
              }}
            >
              {isLoading ? (
                <>
                  <span style={{ display: 'inline-block', animation: 'spin-slow 1s infinite linear' }}>⚡</span>
                  <span>Menjalankan Query (Lazy Load)...</span>
                </>
              ) : (
                <>
                  <span>⚡ Jalankan Query Server</span>
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

const presetBtnStyle: React.CSSProperties = {
  fontSize: '12px',
  fontWeight: '500',
  background: '#1e293b',
  color: '#38bdf8',
  border: '1px solid rgba(56, 189, 248, 0.3)',
  borderRadius: '8px',
  padding: '6px 12px',
  cursor: 'pointer',
  transition: 'all 0.15s ease',
};

const selectFieldStyle: React.CSSProperties = {
  background: '#0f172a',
  border: '1px solid #334155',
  borderRadius: '8px',
  padding: '6px 10px',
  fontSize: '12px',
  color: '#f8fafc',
  outline: 'none',
  maxWidth: '180px',
  cursor: 'pointer',
};

const inputFieldStyle: React.CSSProperties = {
  flex: 1,
  minWidth: '140px',
  background: '#0f172a',
  border: '1px solid #334155',
  borderRadius: '8px',
  padding: '6px 10px',
  fontSize: '12px',
  color: '#f8fafc',
  outline: 'none',
};

