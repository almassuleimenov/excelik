'use client';

import React, { useState } from 'react';

export default function ExcelComparePage() {
  const [file1, setFile1] = useState<File | null>(null);
  const [file2, setFile2] = useState<File | null>(null);
  const [idColumn, setIdColumn] = useState<string>('ID');
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!file1 || !file2) {
      setError('Пожалуйста, выберите оба файла для сравнения.');
      return;
    }

    setLoading(true);
    setError(null);

    const formData = new FormData();
    formData.append('file1', file1);
    formData.append('file2', file2);
    formData.append('id_column', idColumn);

    try {
      // Замени URL на тот, который выдаст Render после деплоя бэкенда
      const backendUrl = 'https://ЗАМЕНИ_НА_URL_ТВОЕГО_БЭКЕНДА_НА_RENDER.onrender.com/api/v1/compare';
      
      const response = await fetch(backendUrl, {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Произошла ошибка при обработке файлов.');
      }

      // Получаем бинарный ответ (Blob) и инициируем скачивание
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `report_${Date.now()}.xlsx`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err: any) {
      setError(err.message || 'Не удалось связаться с сервером бэкенда.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="min-h-screen bg-neutral-950 text-neutral-100 flex flex-col justify-center items-center p-6 antialiased">
      <div className="w-full max-w-lg bg-neutral-900 border border-neutral-800 rounded-xl p-8 shadow-2xl">
        <h1 className="text-2xl font-light tracking-tight text-neutral-50 mb-2 text-center">
          Excel Comparator
        </h1>
        <p className="text-sm text-neutral-400 text-center mb-8 font-light">
          Двусторонний поиск расхождений (Симметрическая разность)
        </p>

        <form onSubmit={handleSubmit} className="space-y-6">
          <div>
            <label className="block text-xs uppercase tracking-wider text-neutral-400 font-medium mb-2">
              Имя колонки идентификатора (ID)
            </label>
            <input
              type="text"
              value={idColumn}
              onChange={(e) => setIdColumn(e.target.value)}
              className="w-full bg-neutral-950 border border-neutral-800 rounded-lg px-4 py-2.5 text-sm text-neutral-200 focus:outline-none focus:border-neutral-600 transition-colors"
              required
            />
          </div>

          <div className="space-y-4">
            <div>
              <label className="block text-xs uppercase tracking-wider text-neutral-400 font-medium mb-2">
                Файл 1 (Основной)
              </label>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={(e) => setFile1(e.target.files?.[0] || null)}
                className="w-full text-sm text-neutral-400 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:uppercase file:tracking-wider file:font-semibold file:bg-neutral-800 file:text-neutral-200 hover:file:bg-neutral-700 file:cursor-pointer transition-all"
              />
            </div>

            <div>
              <label className="block text-xs uppercase tracking-wider text-neutral-400 font-medium mb-2">
                Файл 2 (Для сверки)
              </label>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={(e) => setFile2(e.target.files?.[0] || null)}
                className="w-full text-sm text-neutral-400 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:uppercase file:tracking-wider file:font-semibold file:bg-neutral-800 file:text-neutral-200 hover:file:bg-neutral-700 file:cursor-pointer transition-all"
              />
            </div>
          </div>

          {error && (
            <div className="p-4 bg-red-950/50 border border-red-900/50 rounded-lg text-sm text-red-400 font-light">
              {error}
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            className="w-full bg-neutral-100 text-neutral-950 rounded-lg py-3 text-sm font-medium hover:bg-neutral-200 transition-colors disabled:bg-neutral-800 disabled:text-neutral-500 cursor-pointer disabled:cursor-not-allowed"
          >
            {loading ? 'Обработка данных...' : 'Сравнить таблицы и скачать отчет'}
          </button>
        </form>
      </div>
    </main>
  );
}