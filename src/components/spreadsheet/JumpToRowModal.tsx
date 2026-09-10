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
  const inputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isOpen) {
      setErrorMessage(null);
      setIsJumping(false);
      setTimeout(() => {
        inputRef.current?.focus();
        inputRef.current?.select();
      }, 50);
    }
  }, [isOpen]);

  if (!isOpen) return null;

  const handleExecuteJump = async (rowNum: number) => {
    if (isNaN(rowNum) || rowNum < 1 || rowNum > maxRow) {
      setErrorMessage(`Nomor baris harus antara 1 dan ${maxRow.toLocaleString()}.`);
      return;
    }

    setIsJumping(true);
    setErrorMessage(null);

    try {
      await onJump(rowNum);
      onClose();
    } catch (err: any) {
      setErrorMessage(err?.message || 'Gagal melompat ke baris.');
    } finally {
      setIsJumping(false);
    }
  };

  const onSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const row = parseInt(targetRowStr.replace(/[,._]/g, ''), 10);
    handleExecuteJump(row);
  };

  const presets = [
    { label: 'Baris 1', row: 1 },
    { label: 'Baris 100.000', row: 100000 },
    { label: 'Baris 1.000.000', row: 1000000 },
    { label: 'Baris 4.000.000', row: 4000000 },
    { label: `Baris Terakhir (${maxRow.toLocaleString()})`, row: maxRow },
  ].filter((p) => p.row <= maxRow);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
      <div
        className="w-full max-w-md bg-slate-900 border border-slate-700 rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-150"
        role="dialog"
      >
        <div className="flex items-center justify-between border-b border-slate-800 px-5 py-4">
          <div className="flex items-center gap-2">
            <span className="flex h-7 w-7 items-center justify-center rounded-lg bg-indigo-500/20 text-indigo-400 font-bold text-sm">
              #
            </span>
            <h3 className="font-semibold text-slate-100 text-base">Lompat ke Baris (Ctrl+G)</h3>
          </div>
          <button
            onClick={onClose}
            className="text-slate-400 hover:text-slate-200 transition-colors text-lg"
          >
            ✕
          </button>
        </div>

        <form onSubmit={onSubmit} className="p-5 space-y-4">
          <div>
            <label className="block text-xs font-medium text-slate-300 mb-1.5">
              Masukkan Nomor Baris (1 – {maxRow.toLocaleString()})
            </label>
            <input
              ref={inputRef}
              type="text"
              value={targetRowStr}
              onChange={(e) => setTargetRowStr(e.target.value)}
              placeholder="Contoh: 1000000"
              className="w-full rounded-lg bg-slate-800 border border-slate-700 px-3.5 py-2.5 text-sm text-slate-100 placeholder-slate-500 focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono"
            />
          </div>

          {errorMessage && (
            <p className="text-xs text-rose-400 bg-rose-500/10 border border-rose-500/20 rounded p-2">
              {errorMessage}
            </p>
          )}

          <div>
            <span className="block text-xs font-medium text-slate-400 mb-2">Pintasan Cepat:</span>
            <div className="flex flex-wrap gap-1.5">
              {presets.map((p) => (
                <button
                  key={p.row}
                  type="button"
                  onClick={() => {
                    setTargetRowStr(p.row.toString());
                    handleExecuteJump(p.row);
                  }}
                  className="text-xs bg-slate-800 hover:bg-slate-700 text-slate-300 border border-slate-700/60 rounded px-2.5 py-1 transition-colors"
                >
                  {p.label}
                </button>
              ))}
            </div>
          </div>

          <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-800">
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg px-4 py-2 text-xs font-medium text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-colors"
            >
              Batal
            </button>
            <button
              type="submit"
              disabled={isJumping}
              className="rounded-lg bg-indigo-600 px-4 py-2 text-xs font-medium text-white hover:bg-indigo-500 disabled:opacity-50 transition-colors shadow-sm"
            >
              {isJumping ? 'Melompat...' : 'Lompat ke Baris'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
