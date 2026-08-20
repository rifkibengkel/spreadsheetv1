'use client';

import React, { useEffect, useRef, useState, useCallback } from 'react';
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import '@univerjs/presets/lib/styles/preset-sheets-core.css';
import '@univerjs/preset-sheets-core/lib/index.css';

import { initialWorkbookData } from '@/lib/univer/univerConfig';
import SpreadsheetToolbar from './SpreadsheetToolbar';

export default function Spreadsheet() {
  const containerRef = useRef<HTMLDivElement>(null);
  const univerRef = useRef<any>(null);
  const [univerAPI, setUniverAPI] = useState<any>(null);

  const [activeFileName, setActiveFileName] = useState('workbook-initial.xlsx');
  const [activeFileFormat, setActiveFileFormat] = useState<'xlsx' | 'csv'>('xlsx');

  const handleFileImported = useCallback((filename: string, format: 'xlsx' | 'csv') => {
    setActiveFileName(filename);
    setActiveFileFormat(format);
  }, []);

  useEffect(() => {
    if (!containerRef.current) return;

    if (univerRef.current?.univer) {
      setUniverAPI(univerRef.current.univerAPI);
      return;
    }

    let currentRefObj: any = null;
    let formulaWorkerInstance: Worker | null = null;

    try {
      if (containerRef.current && containerRef.current.children.length > 0) {
        containerRef.current.innerHTML = '';
      }

      if (typeof window !== 'undefined' && window.Worker) {
        formulaWorkerInstance = new Worker(new URL('../../lib/workers/formula.worker.ts', import.meta.url), {
          type: 'module',
        });
      }

      const corePreset = UniverSheetsCorePreset({
        container: containerRef.current,
        workerURL: formulaWorkerInstance as any,
      } as any);

      const { univer, univerAPI: api } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
          [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS),
        },
        presets: [corePreset],
      });

      currentRefObj = { univer, univerAPI: api, worker: formulaWorkerInstance };
      univerRef.current = currentRefObj;
      (window as any).univerAPI = api;
      (window as any).univer = univer;

      // Set initial formula calculation mode to NO_CALCULATION before sheet creation
      try {
        const formulaEngine = api.getFormula ? api.getFormula() : null;
        if (formulaEngine && formulaEngine.setInitialFormulaComputing) {
          formulaEngine.setInitialFormulaComputing(2); // CalculationMode.NO_CALCULATION
        }
      } catch (err) {
        console.warn('Initial formula computing setup warning:', err);
      }

      // Initialize the workbook snapshot
      api.createUniverSheet(initialWorkbookData);
      setUniverAPI(api);

      setTimeout(() => {
        if (typeof window !== 'undefined') {
          window.dispatchEvent(new Event('resize'));
        }
      }, 100);
    } catch (e) {
      console.error('Failed to initialize Univer:', e);
    }

    return () => {
      const activeObj = currentRefObj;

      setTimeout(() => {
        if (univerRef.current !== activeObj && activeObj?.univer) {
          try {
            if (activeObj.worker) {
              activeObj.worker.terminate();
            }
            activeObj.univer.dispose();
          } catch (e) {
            console.warn('Dispose warning:', e);
          }
        }
      }, 300);
    };
  }, []);

  return (
    <div className="univer-wrapper">
      <div className="univer-header">
        <p>Edit, Import, & Save Excel / CSV Files directly | Glory-Glory Masferr.AI</p>

        {/* <p>Edit, Import, & Save Excel / CSV Files directly to Desktop</p> */}
      </div>
      <SpreadsheetToolbar
        univerAPI={univerAPI}
        activeFileName={activeFileName}
        activeFileFormat={activeFileFormat}
        onFileImported={handleFileImported}
      />
      <div className="univer-container-outer">
        <div ref={containerRef} className="univer-container-inner" />
      </div>
    </div>
  );
}

