import React, { useState, useRef } from 'react';
import { importExcelToWorkbookDataAsync, replaceUniverWorkbook } from '@/lib/import/excel';
import { importCSVBatched } from '@/lib/import/csv';

interface ImportModalProps {
  univerAPI: any;
  onClose: () => void;
  onFileImported?: (filename: string, format: 'xlsx' | 'csv') => void;
}

export default function ImportModal({ univerAPI, onClose, onFileImported }: ImportModalProps) {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [statusMessage, setStatusMessage] = useState('');
  const [error, setError] = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      setFile(e.dataTransfer.files[0]);
      setError('');
    }
  };

  const handleImport = async () => {
    if (!file || !univerAPI) return;
    setLoading(true);
    setError('');
    setProgress(0);
    setStatusMessage('Memulai impor file...');

    const ext = file.name.split('.').pop()?.toLowerCase() as 'xlsx' | 'xls' | 'csv' | undefined;

    try {
      if (ext === 'xlsx' || ext === 'xls') {
        const payload = await importExcelToWorkbookDataAsync(file, (percent, msg) => {
          setProgress(percent);
          if (msg) setStatusMessage(msg);
        });

        setProgress(100);
        setStatusMessage('Impor Berhasil! Memuat ke spreadsheet...');
        onFileImported?.(file.name, 'xlsx');

        // Close modal immediately so UI does not get stuck at 100%
        onClose();

        // Defer heavy Univer Sheet creation slightly to allow modal unmount cleanly
        setTimeout(() => {
          try {
            replaceUniverWorkbook(univerAPI, payload.workbookData);
          } catch (e: any) {
            console.error('Replace workbook error:', e);
          }
        }, 50);
      } else if (ext === 'csv') {
        importCSVBatched(
          file,
          univerAPI,
          (percent) => {
            setProgress(percent);
            setStatusMessage(`Membaca CSV (${percent}%)...`);
          },
          () => {
            setStatusMessage('Impor CSV Berhasil!');
            onFileImported?.(file.name, 'csv');
            onClose();
          },
          (err) => {
            setError(err.message || 'CSV Import gagal.');
            setLoading(false);
          }
        );
      } else {
        setError('Format tidak didukung. Harap gunakan file .xlsx, .xls, atau .csv');
        setLoading(false);
      }
    } catch (e: any) {
      console.error('Import Error:', e);
      setError(e.message || 'Proses impor gagal.');
      setLoading(false);
    }
  };

  return (
    <div style={overlayStyle}>
      <div style={modalStyle}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
          <h2 style={{ margin: 0, fontSize: '20px', color: '#f8fafc', fontWeight: '700' }}>
            📥 Import Spreadsheet (Excel / CSV)
          </h2>
          <button onClick={onClose} disabled={loading} style={closeXBtnStyle}>✕</button>
        </div>

        <p style={{ color: '#94a3b8', fontSize: '13px', margin: '0 0 16px 0' }}>
          Pilih file `.xlsx` atau `.csv` (mendukung hingga 2M+ baris tanpa freeze).
        </p>

        {/* Dropzone */}
        <div
          onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={handleFileDrop}
          onClick={() => !loading && fileInputRef.current?.click()}
          style={{
            ...dropzoneStyle,
            borderColor: isDragging ? '#38bdf8' : file ? '#10b981' : '#334155',
            background: isDragging ? 'rgba(56, 189, 248, 0.1)' : file ? 'rgba(16, 185, 129, 0.08)' : '#0f172a',
          }}
        >
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx, .xls, .csv"
            onChange={(e) => { setFile(e.target.files?.[0] || null); setError(''); }}
            style={{ display: 'none' }}
            disabled={loading}
          />

          {file ? (
            <div>
              <div style={{ fontSize: '28px', marginBottom: '6px' }}>📄</div>
              <strong style={{ color: '#f8fafc', fontSize: '14px', display: 'block' }}>{file.name}</strong>
              <span style={{ color: '#94a3b8', fontSize: '12px' }}>
                Ukuran: {(file.size / (1024 * 1024)).toFixed(2)} MB
              </span>
            </div>
          ) : (
            <div>
              <div style={{ fontSize: '28px', marginBottom: '6px' }}>📂</div>
              <strong style={{ color: '#e2e8f0', fontSize: '14px', display: 'block' }}>
                Klik atau Tarik File Excel / CSV ke Sini
              </strong>
              <span style={{ color: '#64748b', fontSize: '12px' }}>Mendukung .xlsx, .xls, .csv</span>
            </div>
          )}
        </div>

        {error && <div style={errorBannerStyle}>{error}</div>}

        {/* Progress Bar UI */}
        {loading && (
          <div style={{ marginTop: '16px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '6px', fontSize: '12px' }}>
              <span style={{ color: '#38bdf8', fontWeight: 'bold' }}>{statusMessage}</span>
              <span style={{ color: '#38bdf8', fontWeight: 'bold' }}>{progress}%</span>
            </div>

            <div style={{ background: '#0f172a', height: '10px', borderRadius: '5px', overflow: 'hidden', border: '1px solid #334155' }}>
              <div
                style={{
                  background: 'linear-gradient(90deg, #38bdf8 0%, #10b981 100%)',
                  width: `${progress}%`,
                  height: '100%',
                  transition: 'width 0.2s linear',
                  boxShadow: '0 0 12px rgba(56, 189, 248, 0.6)',
                }}
              />
            </div>
          </div>
        )}

        <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end', marginTop: '24px' }}>
          <button onClick={onClose} disabled={loading} style={cancelBtnStyle}>
            Batal
          </button>
          <button onClick={handleImport} disabled={!file || loading} style={importBtnStyle(!file || loading)}>
            {loading ? `Memproses (${progress}%)` : '🚀 Mulai Impor'}
          </button>
        </div>
      </div>
    </div>
  );
}

const overlayStyle: React.CSSProperties = {
  position: 'fixed',
  top: 0, left: 0, right: 0, bottom: 0,
  background: 'rgba(15, 23, 42, 0.8)',
  backdropFilter: 'blur(8px)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  zIndex: 1000,
};

const modalStyle: React.CSSProperties = {
  background: '#1e293b',
  border: '1px solid #334155',
  padding: '24px',
  borderRadius: '16px',
  width: '460px',
  boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.7)',
  color: '#f8fafc',
};

const dropzoneStyle: React.CSSProperties = {
  border: '2px dashed #334155',
  borderRadius: '12px',
  padding: '24px',
  textAlign: 'center',
  cursor: 'pointer',
  transition: 'all 0.2s ease-in-out',
};

const closeXBtnStyle: React.CSSProperties = {
  background: 'transparent',
  border: 'none',
  color: '#94a3b8',
  fontSize: '18px',
  cursor: 'pointer',
};

const cancelBtnStyle: React.CSSProperties = {
  padding: '10px 18px',
  borderRadius: '8px',
  border: '1px solid #475569',
  background: 'transparent',
  color: '#cbd5e1',
  cursor: 'pointer',
  fontWeight: 'bold',
  fontSize: '13px',
};

const importBtnStyle = (disabled: boolean): React.CSSProperties => ({
  padding: '10px 20px',
  borderRadius: '8px',
  border: 'none',
  background: disabled ? '#334155' : 'linear-gradient(135deg, #0284c7 0%, #0d9488 100%)',
  color: disabled ? '#94a3b8' : '#ffffff',
  fontWeight: 'bold',
  cursor: disabled ? 'not-allowed' : 'pointer',
  fontSize: '13px',
  boxShadow: disabled ? 'none' : '0 4px 14px rgba(2, 132, 199, 0.4)',
  transition: 'all 0.2s ease-in-out',
});

const errorBannerStyle: React.CSSProperties = {
  marginTop: '12px',
  padding: '10px 14px',
  borderRadius: '8px',
  background: '#7f1d1d',
  color: '#fca5a5',
  fontSize: '13px',
  border: '1px solid #991b1b',
};
