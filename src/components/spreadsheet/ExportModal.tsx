import React, { useState } from 'react';
import { exportWorkbookToExcel } from '@/lib/export/excel';
import { exportActiveSheetToCSV } from '@/lib/export/csv';

interface ExportModalProps {
  univerAPI: any;
  onClose: () => void;
  activeFileName?: string;
  activeFileFormat?: 'xlsx' | 'csv';
}

export default function ExportModal({ univerAPI, onClose, activeFileName, activeFileFormat = 'xlsx' }: ExportModalProps) {
  const [format, setFormat] = useState<'xlsx' | 'csv'>(activeFileFormat);
  const [filenameInput, setFilenameInput] = useState(() => activeFileName || `workbook-export.${activeFileFormat}`);
  const [loading, setLoading] = useState(false);
  const [progressPercent, setProgressPercent] = useState(0);
  const [progressMessage, setProgressMessage] = useState('');
  const [error, setError] = useState('');
  const [successMessage, setSuccessMessage] = useState('');

  // Update extension if user changes format manually
  const handleFormatChange = (newFormat: 'xlsx' | 'csv') => {
    setFormat(newFormat);
    const baseName = filenameInput.substring(0, filenameInput.lastIndexOf('.')) || filenameInput;
    setFilenameInput(`${baseName}.${newFormat}`);
  };

  // const getExportBlob = async (onProgress: (percent: number, msg: string) => void): Promise<{ blob: Blob; targetName: string }> => {
  //   let blob: Blob;
  //   let targetName = filenameInput.trim();
    
  //   if (!targetName.endsWith(`.${format}`)) {
  //     targetName += `.${format}`;
  //   }
    
  //   console.log(targetName,"LUAAR TARGET");
  //   console.log(format,"FORMAAT");
    
  //   if (format === 'xlsx') {
  //     blob = await exportWorkbookToExcel(univerAPI, onProgress);
  //     console.log(blob,"DALEMM");
  //   } else {
  //     blob = await exportActiveSheetToCSV(univerAPI, onProgress);
  //     console.log(blob,"DALEMM2");
  //   }

  //   return { blob, targetName };
  // };

const getExportBlob = async (
  onProgress: (percent: number, msg: string) => void
): Promise<{ blob: Blob; targetName: string }> => {
  let blob: Blob;
  let targetName = filenameInput.trim();

  if (!targetName.endsWith(`.${format}`)) {
    targetName += `.${format}`;
  }

  console.log('[EXPORT MODAL] Target name:', targetName);
  console.log('[EXPORT MODAL] Format:', format);
  console.log('[EXPORT MODAL] univerAPI exists:', !!univerAPI);

  if (format === 'xlsx') {
    console.log('[EXPORT MODAL] BEFORE exportWorkbookToExcel');

    blob = await exportWorkbookToExcel(univerAPI, (percent, msg) => {
      console.log('[EXPORT MODAL] Progress:', percent, msg);
      onProgress(percent, msg);
    });

    console.log('[EXPORT MODAL] AFTER exportWorkbookToExcel');
    console.log('[EXPORT MODAL] Blob:', blob);
    console.log('[EXPORT MODAL] Blob size:', blob?.size);
    console.log('[EXPORT MODAL] Blob type:', blob?.type);
  } else {
    console.log('[EXPORT MODAL] BEFORE exportActiveSheetToCSV');

    blob = await exportActiveSheetToCSV(univerAPI, (percent, msg) => {
      console.log('[EXPORT MODAL] Progress:', percent, msg);
      onProgress(percent, msg);
    });

    console.log('[EXPORT MODAL] AFTER exportActiveSheetToCSV');
    console.log('[EXPORT MODAL] Blob:', blob);
    console.log('[EXPORT MODAL] Blob size:', blob?.size);
    console.log('[EXPORT MODAL] Blob type:', blob?.type);
  }

  console.log('[EXPORT MODAL] Returning blob');

  return { blob, targetName };
};

  const handleDownload = async () => {
    if (!univerAPI) return;
    setLoading(true);
    setError('');
    setSuccessMessage('');
    setProgressPercent(0);
    setProgressMessage('Menyiapkan file untuk diekspor...');

    try {
      const { blob, targetName } = await getExportBlob((percent, msg) => {
        setProgressPercent(percent);
        setProgressMessage(msg);
      });

      setProgressPercent(95);
      setProgressMessage('Mengunduh file ke browser...');

      // Trigger download
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = targetName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      setProgressPercent(100);
      setProgressMessage('✓ Pengunduhan selesai!');
      setTimeout(onClose, 800);
    } catch (e: any) {
      setError(e.message || 'Export failed.');
      setLoading(false);
    }
  };

  const handleSaveToDesktop = async () => {
    if (!univerAPI) return;
    setLoading(true);
    setError('');
    setSuccessMessage('');
    setProgressPercent(0);
    setProgressMessage('Menyiapkan proses simpan data...');

    try {
      const { blob, targetName } = await getExportBlob((percent, msg) => {
        setProgressPercent(percent);
        setProgressMessage(msg);
      });

      setProgressPercent(92);
      setProgressMessage('Mengirim file ke Desktop...');

      const formData = new FormData();
      formData.append('file', blob, targetName);
      formData.append('filename', targetName);

      const res = await fetch('/api/desktop/save', {
        method: 'POST',
        body: formData,
      });

      const data = await res.json();
      if (!res.ok || data.error) {
        throw new Error(data.error || 'Failed to save to Desktop.');
      }

      setProgressPercent(100);
      setProgressMessage('✓ Berhasil disimpan ke Desktop!');
      setSuccessMessage(`Berhasil disimpan ke Desktop: ${data.filePath || targetName}`);
      setLoading(false);
      setTimeout(onClose, 1500);
    } catch (e: any) {
      setError(e.message || 'Save to Desktop failed.');
      setLoading(false);
    }
  };

  return (
    <div id="export-modal-overlay" style={overlayStyle}>
      <div style={modalStyle}>
        <h2 style={{ margin: '0 0 8px 0', fontSize: '20px' }}>Ekspor & Simpan Data</h2>
        <p style={{ color: '#94a3b8', fontSize: '13px', marginBottom: '16px', lineHeight: '1.4' }}>
          Simpan perubahan spreadsheet ke Desktop atau unduh sebagai file Excel / CSV.
        </p>

        <div style={{ marginBottom: '14px' }}>
          <label style={{ display: 'block', marginBottom: '6px', fontSize: '13px', fontWeight: 'bold' }}>Nama File:</label>
          <input
            type="text"
            value={filenameInput}
            onChange={(e) => setFilenameInput(e.target.value)}
            disabled={loading}
            style={inputStyle}
          />
        </div>
        
        <div style={{ marginBottom: '16px' }}>
          <label style={{ display: 'block', marginBottom: '6px', fontSize: '13px', fontWeight: 'bold' }}>Format File:</label>
          <select 
            value={format} 
            onChange={(e) => handleFormatChange(e.target.value as any)}
            style={selectStyle}
            disabled={loading}
          >
            <option value="xlsx">Excel Workbook (.xlsx)</option>
            <option value="csv">Active Sheet as CSV (.csv)</option>
          </select>
        </div>

        {loading && (
          <div style={{ marginBottom: '16px', background: '#0f172a', padding: '12px', borderRadius: '8px', border: '1px solid #334155' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
              <span style={{ color: '#38bdf8', fontSize: '12px', fontWeight: 'bold' }}>Proses Ekspor / Simpan File...</span>
              <span id="export-progress-percent" style={{ color: '#38bdf8', fontSize: '12px', fontWeight: 'bold' }}>{progressPercent}%</span>
            </div>
            <div style={{ width: '100%', height: '8px', background: '#1e293b', borderRadius: '4px', overflow: 'hidden' }}>
              <div 
                style={{ 
                  height: '100%', 
                  width: `${progressPercent}%`, 
                  background: 'linear-gradient(90deg, #38bdf8 0%, #818cf8 100%)',
                  transition: 'width 0.2s ease-out'
                }} 
              />
            </div>
            <p id="export-status-message" style={{ color: '#cbd5e1', fontSize: '11px', marginTop: '6px', margin: '6px 0 0 0', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
              {progressMessage || 'Memproses data...'}
            </p>
          </div>
        )}

        {error && <p style={{ color: '#ef4444', marginBottom: '12px', fontSize: '13px' }}>{error}</p>}
        {successMessage && <p style={{ color: '#10b981', marginBottom: '12px', fontWeight: 'bold', fontSize: '13px' }}>{successMessage}</p>}

        <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end', marginTop: '20px', flexWrap: 'wrap' }}>
          <button onClick={onClose} disabled={loading} style={cancelBtnStyle}>Batal</button>
          <button onClick={handleDownload} disabled={loading} style={exportBtnStyle}>
            📥 Unduh File
          </button>
          {/* <button onClick={handleSaveToDesktop} disabled={loading} style={desktopBtnStyle}>
            💾 Simpan ke Desktop
          </button> */}
        </div>
      </div>
    </div>
  );
}

const overlayStyle: React.CSSProperties = {
    position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
    background: 'rgba(0,0,0,0.65)', backdropFilter: 'blur(4px)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000
};
const modalStyle: React.CSSProperties = {
    background: '#1e293b', padding: '24px', borderRadius: '12px', width: '420px',
    boxShadow: '0 25px 50px -12px rgba(0,0,0,0.5)', color: '#f8fafc', border: '1px solid #334155'
};
const selectStyle = {
    width: '100%', padding: '8px 12px', borderRadius: '6px', background: '#0f172a', border: '1px solid #475569', color: '#f8fafc', fontSize: '13px'
};
const inputStyle = {
    width: '100%', padding: '8px 12px', borderRadius: '6px', background: '#0f172a', border: '1px solid #475569', color: '#f8fafc', boxSizing: 'border-box' as const, fontSize: '13px'
};
const cancelBtnStyle = {
    padding: '8px 16px', borderRadius: '6px', border: '1px solid #475569', background: 'transparent', color: '#cbd5e1', cursor: 'pointer', fontSize: '13px'
};
const exportBtnStyle = {
    padding: '8px 16px', borderRadius: '6px', border: 'none', background: '#10b981', color: '#0f172a', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px'
};
const desktopBtnStyle = {
    padding: '8px 16px', borderRadius: '6px', border: 'none', background: '#38bdf8', color: '#0f172a', fontWeight: 'bold', cursor: 'pointer', fontSize: '13px'
};
