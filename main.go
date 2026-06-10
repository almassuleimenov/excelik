package main

import (
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	port := os.Getenv("PORT")
	if port == "" {
		port = "8080"
	}

	http.HandleFunc("/api/v1/compare", corsMiddleware(compareHandler))

	log.Printf("Сервер запущен на порту %s", port)
	if err := http.ListenAndServe(":"+port, nil); err != nil {
		log.Fatalf("Ошибка запуска сервера: %v", err)
	}
}

// corsMiddleware обрабатывает CORS-запросы для интеграции с фронтендом (Vercel)
func corsMiddleware(next http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}
		next(w, r)
	}
}

func compareHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Метод не поддерживается", http.StatusMethodNotAllowed)
		return
	}

	// Ограничиваем объем буфера для мультипарт-формы (32 МБ в памяти, остальное во временные файлы)
	if err := r.ParseMultipartForm(32 << 20); err != nil {
		http.Error(w, "Ошибка парсинга формы: "+err.Error(), http.StatusBadRequest)
		return
	}

	idColumn := r.FormValue("id_column")
	if idColumn == "" {
		idColumn = "ID"
	}

	file1Header, err := r.FormFile("file1")
	if err != nil {
		http.Error(w, "Файл file1 обязателен", http.StatusBadRequest)
		return
	}
	defer file1Header.Close()

	file2Header, err := r.FormFile("file2")
	if err != nil {
		http.Error(w, "Файл file2 обязателен", http.StatusBadRequest)
		return
	}
	defer file2Header.Close()

	// Инициализируем excelize-читатели напрямую из потоков данных
	f1, err := excelize.OpenReader(file1Header)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 1: "+err.Error(), http.StatusUnprocessableEntity)
		return
	}
	defer f1.Close()

	f2, err := excelize.OpenReader(file2Header)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 2: "+err.Error(), http.StatusUnprocessableEntity)
		return
	}
	defer f2.Close()

	sheet1 := f1.GetSheetName(0)
	sheet2 := f2.GetSheetName(0)

	// ШАГ 1: Сбор ID из Файла 1
	setA, err := buildIDSet(f1, sheet1, idColumn)
	if err != nil {
		http.Error(w, "Ошибка индексации Файла 1: "+err.Error(), http.StatusBadRequest)
		return
	}

	// ШАГ 2: Сбор ID из Файла 2
	setB, err := buildIDSet(f2, sheet2, idColumn)
	if err != nil {
		http.Error(w, "Ошибка индексации Файла 2: "+err.Error(), http.StatusBadRequest)
		return
	}

	// Создаем результирующий файл
	out := excelize.NewFile()
	defer out.Close()

	sheetName1 := "Только в Файле 1"
	sheetName2 := "Только в Файле 2"

	_, _ = out.NewSheet(sheetName1)
	_, _ = out.NewSheet(sheetName2)
	_ = out.DeleteSheet("Sheet1") // Удаляем дефолтный лист

	// ШАГ 3: Выявление расхождений и запись
	if err := writeDiscrepancies(f1, sheet1, idColumn, setB, out, sheetName1); err != nil {
		http.Error(w, "Ошибка генерации расхождений для Файла 1: "+err.Error(), http.StatusInternalServerError)
		return
	}

	if err := writeDiscrepancies(f2, sheet2, idColumn, setA, out, sheetName2); err != nil {
		http.Error(w, "Ошибка генерации расхождений для Файла 2: "+err.Error(), http.StatusInternalServerError)
		return
	}

	// Стриминг готового файла клиенту
	filename := fmt.Sprintf("report_%d.xlsx", time.Now().Unix())
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	if _, err := out.WriteTo(w); err != nil {
		log.Printf("Ошибка отправки отчета клиенту: %v", err)
	}
}

// buildIDSet сканирует файл потоком и собирает уникальные ID
func buildIDSet(f *excelize.File, sheet string, idColumn string) (map[string]struct{}, error) {
	set := make(map[string]struct{})
	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	idIdx := -1
	isFirstRow := true

	for rows.Next() {
		cols, err := rows.Columns()
		if err != nil {
			return nil, err
		}

		if isFirstRow {
			for i, col := range cols {
				if strings.TrimSpace(strings.ToLower(col)) == strings.TrimSpace(strings.ToLower(idColumn)) {
					idIdx = i
					break
				}
			}
			if idIdx == -1 {
				return nil, fmt.Errorf("колонка '%s' не найдена", idColumn)
			}
			isFirstRow = false
			continue
		}

		if idIdx < len(cols) {
			val := strings.TrimSpace(cols[idIdx])
			if val != "" {
				set[val] = struct{}{}
			}
		}
	}
	return set, nil
}

// writeDiscrepancies находит строки, ID которых нет в целевой мапе, и записывает их на нужный лист
func writeDiscrepancies(fIn *excelize.File, sheetIn string, idColumn string, targetSet map[string]struct{}, fOut *excelize.File, sheetOut string) error {
	rows, err := fIn.Rows(sheetIn)
	if err != nil {
		return err
	}
	defer rows.Close()

	idIdx := -1
	isFirstRow := true
	outRowIdx := 1

	for rows.Next() {
		cols, err := rows.Columns()
		if err != nil {
			return err
		}

		if isFirstRow {
			// Находим индекс и записываем шапку таблицы
			for i, col := range cols {
				if strings.TrimSpace(strings.ToLower(col)) == strings.TrimSpace(strings.ToLower(idColumn)) {
					idIdx = i
					break
				}
			}
			if err := writeRow(fOut, sheetOut, outRowIdx, cols); err != nil {
				return err
			}
			outRowIdx++
			isFirstRow = false
			continue
		}

		if idIdx < len(cols) {
			val := strings.TrimSpace(cols[idIdx])
			_, exists := targetSet[val]
			
			// Если идентификатора нет в противоположном файле — это расхождение
			if !exists {
				if err := writeRow(fOut, sheetOut, outRowIdx, cols); err != nil {
					return err
				}
				outRowIdx++
			}
		}
	}
	return nil
}

// writeRow вспомогательная функция для безопасной записи строки по ячейкам
func writeRow(f *excelize.File, sheet string, rowIdx int, values []string) error {
	for colIdx, val := range values {
		cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx)
		if err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cell, val); err != nil {
			return err
		}
	}
	return nil
}