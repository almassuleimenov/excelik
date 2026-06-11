package main

import (
	"fmt"
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

func corsMiddleware(next http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		// Для продакшена лучше заменить "*" на "https://tvoy-domen.vercel.app"
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

	if err := r.ParseMultipartForm(32 << 20); err != nil {
		http.Error(w, "Ошибка парсинга формы: "+err.Error(), http.StatusBadRequest)
		return
	}

	idColumn := r.FormValue("id_column")
	if idColumn == "" {
		idColumn = "ID"
	}

	file1, _, err := r.FormFile("file1")
	if err != nil {
		http.Error(w, "Файл file1 обязателен", http.StatusBadRequest)
		return
	}
	defer file1.Close()

	file2, _, err := r.FormFile("file2")
	if err != nil {
		http.Error(w, "Файл file2 обязателен", http.StatusBadRequest)
		return
	}
	defer file2.Close()

	f1, err := excelize.OpenReader(file1)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 1: "+err.Error(), http.StatusUnprocessableEntity)
		return
	}
	defer f1.Close()

	f2, err := excelize.OpenReader(file2)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 2: "+err.Error(), http.StatusUnprocessableEntity)
		return
	}
	defer f2.Close()

	// ШАГ 1: Сбор ID из Файла 1 (со всех листов)
	setA, err := buildIDSet(f1, idColumn)
	if err != nil {
		http.Error(w, "Ошибка индексации Файла 1: "+err.Error(), http.StatusBadRequest)
		return
	}

	// ШАГ 2: Сбор ID из Файла 2 (со всех листов)
	setB, err := buildIDSet(f2, idColumn)
	if err != nil {
		http.Error(w, "Ошибка индексации Файла 2: "+err.Error(), http.StatusBadRequest)
		return
	}

	out := excelize.NewFile()
	defer out.Close()

	sheetName1 := "Только в Файле 1"
	sheetName2 := "Только в Файле 2"

	_, _ = out.NewSheet(sheetName1)
	_, _ = out.NewSheet(sheetName2)
	_ = out.DeleteSheet("Sheet1")

	// ШАГ 3: Потоковая запись расхождений
	if err := writeDiscrepancies(f1, idColumn, setB, out, sheetName1); err != nil {
		http.Error(w, "Ошибка генерации расхождений Файла 1: "+err.Error(), http.StatusInternalServerError)
		return
	}

	if err := writeDiscrepancies(f2, idColumn, setA, out, sheetName2); err != nil {
		http.Error(w, "Ошибка генерации расхождений Файла 2: "+err.Error(), http.StatusInternalServerError)
		return
	}

	filename := fmt.Sprintf("report_%d.xlsx", time.Now().Unix())
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

	if _, err := out.WriteTo(w); err != nil {
		log.Printf("Ошибка отправки отчета клиенту: %v", err)
	}
}

// buildIDSet итерируется по всем листам книги и собирает уникальные ID
func buildIDSet(f *excelize.File, idColumn string) (map[string]struct{}, error) {
	set := make(map[string]struct{})

	for _, sheet := range f.GetSheetList() {
		rows, err := f.Rows(sheet)
		if err != nil {
			continue // Пропускаем проблемные или пустые листы
		}

		idIdx := -1
		isFirstRow := true

		for rows.Next() {
			cols, err := rows.Columns()
			if err != nil {
				break
			}

			if isFirstRow {
				for i, col := range cols {
					if strings.TrimSpace(strings.ToLower(col)) == strings.TrimSpace(strings.ToLower(idColumn)) {
						idIdx = i
						break
					}
				}
				isFirstRow = false
				continue
			}

			if idIdx != -1 && idIdx < len(cols) {
				val := strings.TrimSpace(cols[idIdx])
				if val != "" {
					set[val] = struct{}{}
				}
			}
		}
		rows.Close()
	}
	return set, nil
}

// writeDiscrepancies использует StreamWriter для записи без нагрузки на RAM
func writeDiscrepancies(fIn *excelize.File, idColumn string, targetSet map[string]struct{}, fOut *excelize.File, sheetOut string) error {
	sw, err := fOut.NewStreamWriter(sheetOut)
	if err != nil {
		return err
	}

	outRowIdx := 1
	headerWritten := false

	for _, sheetIn := range fIn.GetSheetList() {
		rows, err := fIn.Rows(sheetIn)
		if err != nil {
			continue
		}

		idIdx := -1
		isFirstRow := true

		for rows.Next() {
			cols, err := rows.Columns()
			if err != nil {
				break
			}

			if isFirstRow {
				// Ищем колонку ID на текущем листе
				for i, col := range cols {
					if strings.TrimSpace(strings.ToLower(col)) == strings.TrimSpace(strings.ToLower(idColumn)) {
						idIdx = i
						break
					}
				}

				// Записываем шапку только один раз для результирующего файла
				if !headerWritten {
					rowVals := make([]interface{}, len(cols))
					for i, v := range cols {
						rowVals[i] = v
					}
					cell, _ := excelize.CoordinatesToCellName(1, outRowIdx)
					if err := sw.SetRow(cell, rowVals); err == nil {
						outRowIdx++
						headerWritten = true
					}
				}
				isFirstRow = false
				continue
			}

			if idIdx != -1 && idIdx < len(cols) {
				val := strings.TrimSpace(cols[idIdx])
				_, exists := targetSet[val]

				if !exists {
					rowVals := make([]interface{}, len(cols))
					for i, v := range cols {
						rowVals[i] = v
					}
					cell, _ := excelize.CoordinatesToCellName(1, outRowIdx)
					_ = sw.SetRow(cell, rowVals)
					outRowIdx++
				}
			}
		}
		rows.Close()
	}

	// Обязательно сбрасываем буфер на диск
	return sw.Flush()
}