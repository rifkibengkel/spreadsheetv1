import React, { useState } from 'react';
import ImportModal from './ImportModal';
import ExportModal from './ExportModal';
import { exportWorkbookToExcel } from '@/lib/export/excel';
import { exportActiveSheetToCSV } from '@/lib/export/csv';

interface SpreadsheetToolbarProps {
  univerAPI: any;
  activeFileName: string;
  activeFileFormat: 'xlsx' | 'csv';
  onFileImported: (filename: string, format: 'xlsx' | 'csv') => void;
}

const SpreadsheetToolbar = React.memo(function SpreadsheetToolbar({
  univerAPI,
  activeFileName,
  activeFileFormat,
  onFileImported,
}: SpreadsheetToolbarProps) {
  const [showImport, setShowImport] = useState(false);
  const [showExport, setShowExport] = useState(false);
  const [savingDesktop, setSavingDesktop] = useState(false);
  const [toastMessage, setToastMessage] = useState<{ text: string; isError?: boolean } | null>(null);

  const showToast = (text: string, isError = false) => {
    setToastMessage({ text, isError });
    setTimeout(() => {
      setToastMessage(null);
    }, 5000);
  };

  const handleSaveToDesktop = async () => {
    if (!univerAPI) {
      showToast('Mesin spreadsheet belum siap.', true);
      return;
    }

    setSavingDesktop(true);

    try {
      let blob: Blob;
      let targetName = activeFileName;

      if (!targetName.endsWith(`.${activeFileFormat}`)) {
        targetName += `.${activeFileFormat}`;
      }

      if (activeFileFormat === 'xlsx') {
        blob = await exportWorkbookToExcel(univerAPI);
      } else {
        blob = await exportActiveSheetToCSV(univerAPI);
      }

      const formData = new FormData();
      formData.append('file', blob, targetName);
      formData.append('filename', targetName);

      const res = await fetch('/api/desktop/save', {
        method: 'POST',
        body: formData,
      });

      const data = await res.json();
      if (!res.ok || data.error) {
        throw new Error(data.error || 'Penyimpanan ke Desktop gagal.');
      }

      const savedName = data.filename || targetName;
      showToast(`✓ Perubahan berhasil disimpan ke Desktop: ${savedName}`);
    } catch (err: any) {
      showToast(err.message || 'Gagal menyimpan file ke Desktop.', true);
    } finally {
      setSavingDesktop(false);
    }
  };

  return (
    <div style={{ marginBottom: '16px' }}>
      <div
        style={{
          display: 'flex',
          gap: '12px',
          padding: '14px 18px',
          background: '#1e293b',
          border: '1px solid #334155',
          borderRadius: '12px',
          alignItems: 'center',
          flexWrap: 'wrap',
          boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.4)',
        }}
      >
        <div style={{ marginRight: 'auto', display: 'flex', alignItems: 'center', gap: '10px' }}>
          <span style={{ fontSize: '18px' }}>📊</span>
          <strong style={{ color: '#f8fafc', fontSize: '15px', fontWeight: '700' }}>
            Workbook Controls
          </strong>
          <span
            style={{
              padding: '5px 12px',
              borderRadius: '20px',
              background: 'rgba(56, 189, 248, 0.1)',
              color: '#38bdf8',
              fontSize: '12px',
              border: '1px solid rgba(56, 189, 248, 0.3)',
              fontWeight: '600',
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
            }}
          >
            <span style={{ width: '6px', height: '6px', borderRadius: '50%', background: '#38bdf8' }} />
            {activeFileName}
          </span>
        </div>

        <button onClick={() => setShowImport(true)} style={importBtnStyle}>
          📥 Import File (Excel/CSV)
        </button>

        <button
          onClick={handleSaveToDesktop}
          disabled={savingDesktop}
          style={quickSaveBtnStyle(savingDesktop)}
          title="Simpan perubahan terbaru langsung ke Desktop Windows"
        >
          {savingDesktop ? '⏳ Menyimpan...' : '💾 Simpan ke Desktop'}
        </button>

        <button onClick={() => setShowExport(true)} style={exportBtnStyle}>
          📤 Ekspor & Save
        </button>
      </div>

      {toastMessage && (
        <div
          style={{
            marginTop: '10px',
            padding: '12px 18px',
            borderRadius: '8px',
            background: toastMessage.isError ? 'rgba(127, 29, 29, 0.95)' : 'rgba(6, 78, 59, 0.95)',
            color: toastMessage.isError ? '#fca5a5' : '#6ee7b7',
            fontSize: '13px',
            fontWeight: 'bold',
            border: toastMessage.isError ? '1px solid #991b1b' : '1px solid #047857',
            boxShadow: '0 4px 14px rgba(0, 0, 0, 0.4)',
          }}
        >
          {toastMessage.text}
        </div>
      )}

      {showImport && (
        <ImportModal
          univerAPI={univerAPI}
          onClose={() => setShowImport(false)}
          onFileImported={onFileImported}
        />
      )}

      {showExport && (
        <ExportModal
          univerAPI={univerAPI}
          onClose={() => setShowExport(false)}
          activeFileName={activeFileName}
          activeFileFormat={activeFileFormat}
        />
      )}
    </div>
  );
});

export default SpreadsheetToolbar;

const importBtnStyle: React.CSSProperties = {
  padding: '9px 16px',
  background: 'linear-gradient(135deg, #0284c7 0%, #0369a1 100%)',
  color: '#ffffff',
  border: 'none',
  borderRadius: '8px',
  cursor: 'pointer',
  fontWeight: 'bold',
  fontSize: '13px',
  boxShadow: '0 4px 12px rgba(2, 132, 199, 0.3)',
  transition: 'all 0.2s ease',
};

const quickSaveBtnStyle = (disabled: boolean): React.CSSProperties => ({
  padding: '9px 16px',
  background: disabled ? '#334155' : 'linear-gradient(135deg, #059669 0%, #047857 100%)',
  color: disabled ? '#94a3b8' : '#ffffff',
  border: 'none',
  borderRadius: '8px',
  cursor: disabled ? 'not-allowed' : 'pointer',
  fontWeight: 'bold',
  fontSize: '13px',
  boxShadow: disabled ? 'none' : '0 4px 14px rgba(5, 150, 105, 0.3)',
  transition: 'all 0.2s ease',
});

const exportBtnStyle: React.CSSProperties = {
  padding: '9px 16px',
  background: 'linear-gradient(135deg, #4f46e5 0%, #4338ca 100%)',
  color: '#ffffff',
  border: 'none',
  borderRadius: '8px',
  cursor: 'pointer',
  fontWeight: 'bold',
  fontSize: '13px',
  boxShadow: '0 4px 12px rgba(79, 70, 229, 0.3)',
  transition: 'all 0.2s ease',
};
