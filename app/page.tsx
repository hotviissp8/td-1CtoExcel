'use client';

import { useState, useRef, useCallback, useEffect } from 'react';

// ---- Types ----

interface UploadedFile {
  id: string;
  file: File;
}

interface OutputFile {
  id: string;
  name: string;
  url: string;
  rowCount: number;
  warningCount: number;
}

type ToastType = 'success' | 'error' | 'warning' | 'info';

interface Toast {
  id: string;
  type: ToastType;
  message: string;
}

// ---- Toast icons ----

function ToastIcon({ type }: { type: ToastType }) {
  if (type === 'success') return (
    <svg className="w-5 h-5 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
    </svg>
  );
  if (type === 'error') return (
    <svg className="w-5 h-5 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
    </svg>
  );
  if (type === 'warning') return (
    <svg className="w-5 h-5 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" />
    </svg>
  );
  return (
    <svg className="w-5 h-5 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
    </svg>
  );
}

const TOAST_STYLES: Record<ToastType, string> = {
  success: 'bg-green-800 border border-green-600 text-green-100',
  error:   'bg-red-900 border border-red-600 text-red-100',
  warning: 'bg-yellow-800 border border-yellow-500 text-yellow-100',
  info:    'bg-blue-900 border border-blue-600 text-blue-100',
};

// ---- Spinner ----

function Spinner() {
  return (
    <svg
      className="animate-spin h-5 w-5"
      xmlns="http://www.w3.org/2000/svg"
      fill="none"
      viewBox="0 0 24 24"
    >
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z" />
    </svg>
  );
}

// ---- Main component ----

export default function Home() {
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [outputFiles, setOutputFiles] = useState<OutputFile[]>([]);
  const [processing, setProcessing] = useState(false);
  const [currentFile, setCurrentFile] = useState<string | null>(null);
  const [toasts, setToasts] = useState<Toast[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const dragCounterRef = useRef(0);

  // Auto-dismiss toasts after 5s
  useEffect(() => {
    if (toasts.length === 0) return;
    const timer = setTimeout(() => {
      setToasts(prev => prev.slice(1));
    }, 5000);
    return () => clearTimeout(timer);
  }, [toasts]);

  const addToast = useCallback((type: ToastType, message: string) => {
    const id = Math.random().toString(36).slice(2);
    setToasts(prev => [...prev, { id, type, message }]);
  }, []);

  const removeToast = useCallback((id: string) => {
    setToasts(prev => prev.filter(t => t.id !== id));
  }, []);

  // ---- File selection ----

  const addFiles = useCallback((files: FileList | File[]) => {
    const arr = Array.from(files).filter(f =>
      /\.xlsx?$/i.test(f.name)
    );
    if (arr.length < Array.from(files).length) {
      addToast('warning', 'Некоторые файлы пропущены — принимаются только .xls и .xlsx');
    }
    if (arr.length === 0) return;
    setUploadedFiles(prev => {
      const existingNames = new Set(prev.map(u => u.file.name));
      const newEntries = arr
        .filter(f => !existingNames.has(f.name))
        .map(f => ({ id: Math.random().toString(36).slice(2), file: f }));
      if (newEntries.length < arr.length) {
        addToast('info', 'Дублирующиеся файлы не добавлены');
      }
      return [...prev, ...newEntries];
    });
  }, [addToast]);

  const removeFile = useCallback((id: string) => {
    setUploadedFiles(prev => prev.filter(f => f.id !== id));
  }, []);

  // ---- Drag & drop ----

  const handleDragEnter = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    dragCounterRef.current += 1;
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    dragCounterRef.current -= 1;
    if (dragCounterRef.current === 0) setIsDragging(false);
  }, []);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    dragCounterRef.current = 0;
    setIsDragging(false);
    if (e.dataTransfer.files) addFiles(e.dataTransfer.files);
  }, [addFiles]);

  const handleInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) addFiles(e.target.files);
    e.target.value = '';
  }, [addFiles]);

  // ---- Transform ----

  const handleGo = useCallback(async () => {
    if (uploadedFiles.length === 0 || processing) return;
    setProcessing(true);
    setOutputFiles([]);

    for (const { file } of uploadedFiles) {
      setCurrentFile(file.name);
      try {
        // Processed locally in the browser: no upload, so no server body-size limit (413).
        const { transformExcel } = await import('@/lib/transform');
        const { buffer, warnings, rowCount } = await transformExcel(await file.arrayBuffer());

        const blob = new Blob([buffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const url = URL.createObjectURL(blob);
        const outName = file.name.replace(/\.xlsx?$/i, '') + '_result.xlsx';

        setOutputFiles(prev => [
          ...prev,
          {
            id: Math.random().toString(36).slice(2),
            name: outName,
            url,
            rowCount,
            warningCount: warnings.length,
          },
        ]);

        addToast('success', `${outName} — ${rowCount} строк обработано`);
        for (const w of warnings) addToast('warning', w);
      } catch (err) {
        addToast('error', `${file.name}: ${err instanceof Error ? err.message : 'Неизвестная ошибка'}`);
      }
    }

    setCurrentFile(null);
    setProcessing(false);
  }, [uploadedFiles, processing, addToast]);

  // ---- Download helper ----

  const downloadFile = useCallback((url: string, name: string) => {
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    a.click();
  }, []);

  const downloadAll = useCallback(() => {
    outputFiles.forEach(({ url, name }) => downloadFile(url, name));
  }, [outputFiles, downloadFile]);

  const reset = useCallback(() => {
    outputFiles.forEach(({ url }) => URL.revokeObjectURL(url));
    setUploadedFiles([]);
    setOutputFiles([]);
    setToasts([]);
  }, [outputFiles]);

  // ---- Render ----

  const canGo = uploadedFiles.length > 0 && !processing;

  return (
    <div className="min-h-screen bg-slate-900 text-slate-100 p-6 flex flex-col">
      {/* Header */}
      <header className="text-center mb-8">
        <h1 className="text-4xl font-bold text-white tracking-tight">ТД 1С - Excel</h1>
      </header>

      {/* Three-column layout */}
      <main className="flex-1 max-w-6xl mx-auto w-full flex gap-6 items-start">

        {/* LEFT — Drop zone */}
        <section className="flex-1 flex flex-col gap-4">
          <h2 className="text-sm font-semibold uppercase tracking-wider text-slate-400">
            Исходные файлы
          </h2>

          {/* Drop zone */}
          <div
            onClick={() => !processing && fileInputRef.current?.click()}
            onDragEnter={handleDragEnter}
            onDragLeave={handleDragLeave}
            onDragOver={handleDragOver}
            onDrop={handleDrop}
            className={[
              'relative rounded-xl border-2 border-dashed transition-all duration-200 cursor-pointer select-none',
              'min-h-[160px] flex flex-col items-center justify-center p-6 text-center',
              isDragging
                ? 'border-blue-400 bg-blue-900/30'
                : 'border-slate-600 bg-slate-800/60 hover:border-slate-400 hover:bg-slate-800',
              processing ? 'pointer-events-none opacity-60' : '',
            ].join(' ')}
          >
            <svg className="w-10 h-10 mb-3 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5}
                d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
            </svg>
            {isDragging ? (
              <p className="text-blue-300 font-medium">Отпустите файлы здесь</p>
            ) : (
              <>
                <p className="text-slate-300 font-medium">Перетащите .xlsx файлы</p>
                <p className="text-slate-500 text-sm mt-1">или нажмите для выбора</p>
              </>
            )}
            <input
              ref={fileInputRef}
              type="file"
              accept=".xls,.xlsx"
              multiple
              className="hidden"
              onChange={handleInputChange}
            />
          </div>

          {/* File list */}
          {uploadedFiles.length > 0 && (
            <ul className="flex flex-col gap-2">
              {uploadedFiles.map(({ id, file }) => (
                <li
                  key={id}
                  className={[
                    'flex items-center justify-between gap-3 px-4 py-3',
                    'rounded-lg bg-slate-800 border border-slate-700',
                    processing && currentFile === file.name ? 'border-blue-500 bg-blue-900/20' : '',
                  ].join(' ')}
                >
                  <div className="flex items-center gap-2 min-w-0">
                    {processing && currentFile === file.name && (
                      <span className="text-blue-400"><Spinner /></span>
                    )}
                    <svg className="w-5 h-5 text-green-400 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                        d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                    </svg>
                    <span className="text-sm text-slate-200 truncate" title={file.name}>{file.name}</span>
                  </div>
                  <button
                    onClick={e => { e.stopPropagation(); removeFile(id); }}
                    disabled={processing}
                    className="text-slate-500 hover:text-red-400 transition-colors disabled:opacity-40 flex-shrink-0"
                    aria-label="Удалить файл"
                  >
                    <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                    </svg>
                  </button>
                </li>
              ))}
            </ul>
          )}
        </section>

        {/* CENTER — GO button */}
        <section className="flex flex-col items-center justify-center pt-12 w-36 flex-shrink-0 gap-3">
          <button
            onClick={handleGo}
            disabled={!canGo}
            className={[
              'w-28 h-28 rounded-full text-white font-bold text-2xl',
              'flex flex-col items-center justify-center gap-2',
              'transition-all duration-200 shadow-lg',
              canGo
                ? 'bg-blue-600 hover:bg-blue-500 hover:scale-105 active:scale-95 shadow-blue-900/50'
                : 'bg-slate-700 cursor-not-allowed opacity-60',
            ].join(' ')}
          >
            {processing ? (
              <>
                <Spinner />
                <span className="text-xs font-medium leading-tight text-center">Processing...</span>
              </>
            ) : (
              'GO'
            )}
          </button>
          {uploadedFiles.length > 0 && !processing && (
            <p className="text-xs text-slate-500 text-center">
              {uploadedFiles.length} файл{uploadedFiles.length > 1 ? 'а' : ''}
            </p>
          )}
          {(uploadedFiles.length > 0 || outputFiles.length > 0) && !processing && (
            <button
              onClick={reset}
              className="text-xs text-slate-500 hover:text-red-400 transition-colors underline underline-offset-2"
            >
              Начать сначала
            </button>
          )}
        </section>

        {/* RIGHT — Output files */}
        <section className="flex-1 flex flex-col gap-4">
          <div className="flex items-center justify-between">
            <h2 className="text-sm font-semibold uppercase tracking-wider text-slate-400">
              Результаты
            </h2>
            {outputFiles.length > 1 && (
              <button
                onClick={downloadAll}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-700 hover:bg-slate-600 transition-colors text-slate-200 text-xs font-medium"
              >
                <svg className="w-3.5 h-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                </svg>
                Скачать все
              </button>
            )}
          </div>

          {outputFiles.length === 0 ? (
            <div className="rounded-xl border-2 border-dashed border-slate-700 bg-slate-800/30 min-h-[160px] flex items-center justify-center">
              <p className="text-slate-600 text-sm">Результаты появятся здесь</p>
            </div>
          ) : (
            <ul className="flex flex-col gap-2">
              {outputFiles.map(({ id, name, url, rowCount, warningCount }) => (
                <li
                  key={id}
                  className="flex items-center justify-between gap-3 px-4 py-3 rounded-lg bg-slate-800 border border-slate-700"
                >
                  <div className="flex items-center gap-2 min-w-0">
                    <svg className="w-5 h-5 text-blue-400 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                        d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                    </svg>
                    <div className="min-w-0">
                      <p className="text-sm text-slate-200 truncate" title={name}>{name}</p>
                      <p className="text-xs text-slate-500 mt-0.5">
                        {rowCount} строк
                        {warningCount > 0 && (
                          <span className="ml-2 inline-flex items-center gap-1 px-1.5 py-0.5 rounded bg-yellow-800/60 text-yellow-300 text-xs">
                            <svg className="w-3 h-3" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                                d="M12 9v2m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z" />
                            </svg>
                            {warningCount}
                          </span>
                        )}
                      </p>
                    </div>
                  </div>
                  <button
                    onClick={() => downloadFile(url, name)}
                    className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 transition-colors text-white text-sm font-medium flex-shrink-0"
                  >
                    <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                    </svg>
                    Скачать
                  </button>
                </li>
              ))}
            </ul>
          )}
        </section>
      </main>

      {/* Toast notifications — bottom-right */}
      <div
        aria-live="polite"
        className="fixed bottom-6 right-6 flex flex-col gap-2 z-50 max-w-sm w-full pointer-events-none"
      >
        {toasts.map(toast => (
          <div
            key={toast.id}
            className={[
              'flex items-start gap-3 px-4 py-3 rounded-xl shadow-xl pointer-events-auto',
              'transition-all',
              TOAST_STYLES[toast.type],
            ].join(' ')}
          >
            <ToastIcon type={toast.type} />
            <p className="text-sm flex-1 leading-snug">{toast.message}</p>
            <button
              onClick={() => removeToast(toast.id)}
              className="opacity-60 hover:opacity-100 transition-opacity flex-shrink-0 mt-0.5"
              aria-label="Закрыть"
            >
              <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
              </svg>
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}
