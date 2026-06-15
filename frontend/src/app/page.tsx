'use client';
import React, { useState, useEffect } from 'react';

const formatFileSize = (bytes: number): string => {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
};

export default function ExcelAppPage() {
  const [mode, setMode] = useState<'compare' | 'enrich'>('compare');

  const [file1, setFile1] = useState<File | null>(null);
  const [file2, setFile2] = useState<File | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [elapsedTime, setElapsedTime] = useState<number>(0);

  const [idColumn, setIdColumn] = useState<string>('ID');
  const [matchColumns, setMatchColumns] = useState<string>('ФИО, ИИН, Период услуги, Услуга, Сумма');
  const [targetColumn, setTargetColumn] = useState<string>('ID услуги');

  useEffect(() => {
    let interval: ReturnType<typeof setInterval>;
    if (loading) {
      interval = setInterval(() => {
        setElapsedTime((prev) => prev + 1);
      }, 1000);
    } else {
      setElapsedTime(0);
    }
    return () => clearInterval(interval);
  }, [loading]);

  const handleFileChange = (
    e: React.ChangeEvent<HTMLInputElement>,
    setFile: React.Dispatch<React.SetStateAction<File | null>>
  ) => {
    setFile(e.target.files?.[0] || null);
    setError(null);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    if (!file1 || !file2) {
      setError('Пожалуйста, загрузите оба файла.');
      return;
    }

    setLoading(true);
    setError(null);
    setElapsedTime(0);

    const formData = new FormData();
    formData.append('file1', file1);
    formData.append('file2', file2);

    if (mode === 'compare') {
      if (!idColumn.trim()) return setError('Имя колонки идентификатора не может быть пустым.');
      formData.append('id_column', idColumn.trim());
    } else {
      if (!matchColumns.trim() || !targetColumn.trim()) {
        setError('Поля колонок для сверки и целевой колонки обязательны.');
        setLoading(false);
        return;
      }
      formData.append('match_columns', matchColumns.trim());
      formData.append('target_column', targetColumn.trim());
    }

    try {
      const baseUrl = process.env.NEXT_PUBLIC_API_URL || 'https://excelik.onrender.com';
      const endpoint = mode === 'compare' ? '/api/v1/compare' : '/api/v1/enrich';
      
      const response = await fetch(`${baseUrl}${endpoint}`, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        let errorMessage = 'Произошла ошибка при обработке файлов на сервере.';
        try {
          const errorData = await response.json();
          errorMessage = errorData.detail || errorData.error || errorMessage;
        } catch {
          errorMessage = `Ошибка сервера: ${response.status} ${response.statusText} (Возможно, сервер ушел в sleep-мод или таймаут)`;
        }
        throw new Error(errorMessage);
      }

      const blob = await response.blob();
      if (blob.size === 0) throw new Error('Сервер вернул пустой файл.');

      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = mode === 'compare' 
        ? `Отчет_сверка_${new Date().toISOString().slice(0, 10)}.xlsx`
        : `Отчет_обогащение_${new Date().toISOString().slice(0, 10)}.xlsx`;
      
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

    } catch (err: unknown) {
      if (err instanceof TypeError && err.message === 'Failed to fetch') {
        setError('Не удалось связаться с сервером. Дождитесь выхода Render из спящего режима (до 50 секунд) и повторите.');
      } else if (err instanceof Error) {
        setError(err.message);
      } else {
        setError('Неизвестная системная ошибка.');
      }
    } finally {
      setLoading(false);
    }
  };

  const FileUploadZone = ({ label, file, setFile }: any) => (
    <div className="flex flex-col gap-2">
      <span className="text-xs uppercase tracking-wider text-neutral-400 font-medium">{label}</span>
      {!file ? (
        <label className="flex justify-center w-full h-24 px-4 transition border-2 border-neutral-800 border-dashed rounded-lg appearance-none cursor-pointer hover:border-neutral-600 hover:bg-neutral-900/50 focus:outline-none">
          <span className="flex items-center space-x-2 text-neutral-400">
            <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth="2">
              <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
            </svg>
            <span className="text-sm font-light">Нажмите для выбора файла</span>
          </span>
          <input type="file" className="hidden" accept=".xlsx, .xls" onChange={(e) => handleFileChange(e, setFile)} />
        </label>
      ) : (
        <div className="flex items-center justify-between p-4 border border-neutral-700 bg-neutral-800/30 rounded-lg">
          <div className="flex items-center space-x-3 overflow-hidden">
            <svg xmlns="http://www.w3.org/2000/svg" className="w-6 h-6 text-emerald-500 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
            <div className="flex flex-col truncate">
              <span className="text-sm text-neutral-200 truncate">{file.name}</span>
              <span className="text-xs text-neutral-500">{formatFileSize(file.size)}</span>
            </div>
          </div>
          <button type="button" onClick={() => setFile(null)} className="p-2 text-neutral-500 hover:text-red-400 transition-colors">
            <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5" viewBox="0 0 20 20" fill="currentColor">
              <path fillRule="evenodd" d="M4.293 4.293a1 1 0 011.414 0L10 8.586l4.293-4.293a1 1 0 111.414 1.414L11.414 10l4.293 4.293a1 1 0 01-1.414 1.414L10 11.414l-4.293 4.293a1 1 0 01-1.414-1.414L8.586 10 4.293 5.707a1 1 0 010-1.414z" clipRule="evenodd" />
            </svg>
          </button>
        </div>
      )}
    </div>
  );

  return (
    <main className="min-h-screen bg-neutral-950 text-neutral-100 flex flex-col justify-center items-center p-6 antialiased selection:bg-emerald-500/30">
      <div className="w-full max-w-xl bg-neutral-900 border border-neutral-800 rounded-2xl p-8 shadow-2xl">
        
        {/* Навигация (Табы) */}
        <div className="flex space-x-1 bg-neutral-950/50 p-1 rounded-lg mb-8 border border-neutral-800">
          <button
            onClick={() => { setMode('compare'); setError(null); }}
            className={`flex-1 py-2 text-sm font-medium rounded-md transition-all ${mode === 'compare' ? 'bg-neutral-800 text-white shadow-sm' : 'text-neutral-500 hover:text-neutral-300'}`}
          >
            Сверка (Потеряшки)
          </button>
          <button
            onClick={() => { setMode('enrich'); setError(null); }}
            className={`flex-1 py-2 text-sm font-medium rounded-md transition-all ${mode === 'enrich' ? 'bg-neutral-800 text-white shadow-sm' : 'text-neutral-500 hover:text-neutral-300'}`}
          >
            Обогащение (ВПР ID)
          </button>
        </div>

        <form onSubmit={handleSubmit} className="space-y-6">
          
          {mode === 'compare' ? (
            <div className="bg-neutral-950/50 p-5 rounded-xl border border-neutral-800/80">
              <label className="block text-xs uppercase tracking-wider text-neutral-400 font-medium mb-3">Ключ синхронизации (ID)</label>
              <input
                type="text" value={idColumn} onChange={(e) => setIdColumn(e.target.value)}
                className="w-full bg-neutral-900 border border-neutral-700 rounded-lg px-4 py-3 text-sm text-neutral-200 focus:outline-none focus:border-emerald-500/50 transition-all" required
              />
            </div>
          ) : (
            <div className="bg-neutral-950/50 p-5 rounded-xl border border-neutral-800/80 space-y-4">
              <div>
                <label className="block text-xs uppercase tracking-wider text-neutral-400 font-medium mb-2">Колонки для поиска совпадений (через запятую)</label>
                <input
                  type="text" value={matchColumns} onChange={(e) => setMatchColumns(e.target.value)}
                  className="w-full bg-neutral-900 border border-neutral-700 rounded-lg px-4 py-3 text-sm text-neutral-200 focus:outline-none focus:border-emerald-500/50 transition-all" required
                />
              </div>
              <div>
                <label className="block text-xs uppercase tracking-wider text-neutral-400 font-medium mb-2">Название искомой колонки (которую нужно вставить)</label>
                <input
                  type="text" value={targetColumn} onChange={(e) => setTargetColumn(e.target.value)}
                  className="w-full bg-neutral-900 border border-neutral-700 rounded-lg px-4 py-3 text-sm text-neutral-200 focus:outline-none focus:border-emerald-500/50 transition-all" required
                />
              </div>
            </div>
          )}

          <div className="space-y-5">
            <FileUploadZone 
              label={mode === 'compare' ? "Файл 1 (Основной источник)" : "Файл 1 (Тапсырма / Куда вставлять ID)"} 
              file={file1} setFile={setFile1} 
            />
            <FileUploadZone 
              label={mode === 'compare' ? "Файл 2 (Таблица для сверки)" : "Файл 2 (База данных / Откуда брать ID)"} 
              file={file2} setFile={setFile2} 
            />
          </div>

          {error && (
            <div className="p-4 bg-red-950/40 border border-red-900/50 rounded-xl flex items-start gap-3">
              <svg xmlns="http://www.w3.org/2000/svg" className="w-5 h-5 text-red-500 shrink-0 mt-0.5" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7 4a1 1 0 11-2 0 1 1 0 012 0zm-1-9a1 1 0 00-1 1v4a1 1 0 102 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
              </svg>
              <p className="text-sm text-red-400 font-light leading-relaxed">{error}</p>
            </div>
          )}

          <button
            type="submit" disabled={loading || !file1 || !file2}
            className="group relative w-full flex justify-center items-center bg-neutral-100 text-neutral-950 rounded-xl py-3.5 text-sm font-semibold hover:bg-white transition-all disabled:bg-neutral-800 disabled:text-neutral-500 cursor-pointer disabled:cursor-not-allowed"
          >
            {loading ? (
              <div className="flex items-center space-x-2">
                <svg className="animate-spin h-5 w-5 text-neutral-500" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
                <span>Формирование отчета... ({elapsedTime} сек)</span>
              </div>
            ) : (
              mode === 'compare' ? 'Сравнить и скачать отчет' : 'Найти ID и скачать файл'
            )}
          </button>
        </form>
      </div>
    </main>
  );
}