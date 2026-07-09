package main

import (
	"fmt"
	"log"
	"net/http"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

func enrichHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Метод не поддерживается", http.StatusMethodNotAllowed)
		return
	}

	r.Body = http.MaxBytesReader(w, r.Body, maxUploadSize)
	if err := r.ParseMultipartForm(15 << 20); err != nil {
		http.Error(w, "Ошибка формы (возможно превышен лимит 50МБ): "+err.Error(), http.StatusBadRequest)
		return
	}

	matchColsRaw := r.FormValue("match_columns")
	targetColRaw := r.FormValue("target_column")
	if matchColsRaw == "" || targetColRaw == "" {
		http.Error(w, "Необходимы параметры match_columns и target_column", http.StatusBadRequest)
		return
	}

	var matchCols []string
	for _, c := range strings.Split(matchColsRaw, ",") {
		cleanCol := sanitizeKey(c)
		if cleanCol != "" {
			matchCols = append(matchCols, cleanCol)
		}
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

	tempPath1, err := saveToTempFile(file1, "file1")
	if err != nil {
		http.Error(w, "Ошибка сохранения Файла 1", http.StatusInternalServerError)
		return
	}
	defer os.Remove(tempPath1)

	tempPath2, err := saveToTempFile(file2, "file2")
	if err != nil {
		http.Error(w, "Ошибка сохранения Файла 2", http.StatusInternalServerError)
		return
	}
	defer os.Remove(tempPath2)

	f1, err := excelize.OpenFile(tempPath1)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 1", http.StatusUnprocessableEntity)
		return
	}
	defer f1.Close()

	f2, err := excelize.OpenFile(tempPath2)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 2", http.StatusUnprocessableEntity)
		return
	}
	defer f2.Close()

	enrichMap, err := buildEnrichMapStream(f2, matchCols, sanitizeKey(targetColRaw))
	if err != nil {
		http.Error(w, "Ошибка построения базы: "+err.Error(), http.StatusBadRequest)
		return
	}

	out := excelize.NewFile()
	defer out.Close()
	_ = out.DeleteSheet("Sheet1")

	if err := processEnrichStream(f1, enrichMap, matchCols, targetColRaw, out); err != nil {
		http.Error(w, "Ошибка обработки данных: "+err.Error(), http.StatusInternalServerError)
		return
	}

	tempOutPath, err := createSafeTempOutFile("enriched_report")
	if err != nil {
		http.Error(w, "Ошибка создания отчета", http.StatusInternalServerError)
		return
	}
	defer os.Remove(tempOutPath)

	if err := out.SaveAs(tempOutPath); err != nil {
		log.Printf("Ошибка сохранения отчета: %v", err)
		http.Error(w, "Ошибка генерации файла", http.StatusInternalServerError)
		return
	}

	w.Header().Set("Content-Disposition", "attachment; filename=enriched_report.xlsx")
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	http.ServeFile(w, r, tempOutPath)
}

func buildEnrichMapStream(f *excelize.File, matchCols []string, targetCol string) (map[string]string, error) {
	enrichMap := make(map[string]string)
	foundValidSheet := false

	for _, sheet := range f.GetSheetList() {
		rows, err := f.Rows(sheet)
		if err != nil {
			continue
		}

		matchIdxs := make([]int, len(matchCols))
		for i := range matchIdxs {
			matchIdxs[i] = -1
		}
		targetIdx := -1
		headerFound := false

		for rows.Next() {
			cols, err := rows.Columns()
			if err != nil {
				continue
			}

			if !headerFound {
				for i, colName := range cols {
					normCol := sanitizeKey(colName)
					for j, matchCol := range matchCols {
						if normCol == matchCol {
							matchIdxs[j] = i
						}
					}
					if normCol == targetCol {
						targetIdx = i
					}
				}

				isValid := targetIdx != -1
				for _, idx := range matchIdxs {
					if idx == -1 {
						isValid = false
						break
					}
				}

				if isValid {
					headerFound = true
					foundValidSheet = true
				}
				continue
			}

			if targetIdx != -1 && targetIdx < len(cols) {
				var keyParts []string
				emptyKey := true

				for _, idx := range matchIdxs {
					if idx != -1 && idx < len(cols) {
						normalizedVal := normalizeData(cols[idx])
						keyParts = append(keyParts, normalizedVal)
						if normalizedVal != "" {
							emptyKey = false
						}
					} else {
						keyParts = append(keyParts, "")
					}
				}

				if emptyKey {
					continue
				}

				key := strings.Join(keyParts, "|||")
				if val := strings.TrimSpace(cols[targetIdx]); val != "" {
					enrichMap[key] = val
				}
			}
		}
		rows.Close() // Предотвращаем утечку памяти
	}

	if !foundValidSheet {
		return nil, fmt.Errorf("требуемые колонки для сверки или целевая колонка не найдены в файле базы данных")
	}
	return enrichMap, nil
}

func processEnrichStream(fIn *excelize.File, enrichMap map[string]string, matchCols []string, targetColRaw string, fOut *excelize.File) error {
	foundAnyValidSheet := false

	for _, sheet := range fIn.GetSheetList() {
		_, _ = fOut.NewSheet(sheet)
		sw, err := fOut.NewStreamWriter(sheet)
		if err != nil {
			continue
		}

		rows, err := fIn.Rows(sheet)
		if err != nil {
			continue
		}

		matchIdxs := make([]int, len(matchCols))
		for i := range matchIdxs {
			matchIdxs[i] = -1
		}
		headerFound := false
		outRowIdx := 1

		for rows.Next() {
			cols, err := rows.Columns()
			if err != nil {
				continue
			}

			if !headerFound {
				for i, colName := range cols {
					normCol := sanitizeKey(colName)
					for j, matchCol := range matchCols {
						if normCol == matchCol {
							matchIdxs[j] = i
						}
					}
				}

				isValid := true
				for _, idx := range matchIdxs {
					if idx == -1 {
						isValid = false
						break
					}
				}

				if isValid {
					headerFound = true
					foundAnyValidSheet = true

					rowVals := make([]interface{}, len(cols)+1)
					for i, col := range cols {
						rowVals[i] = col
					}
					rowVals[len(cols)] = "Найденный " + targetColRaw

					cell, _ := excelize.CoordinatesToCellName(1, outRowIdx)
					if err := sw.SetRow(cell, rowVals); err != nil {
						log.Printf("Ошибка записи заголовка: %v", err)
					}
					outRowIdx++
				}
				continue
			}

			// Обработка данных
			rowVals := make([]interface{}, len(cols)+1)
			for i, col := range cols {
				rowVals[i] = col
			}

			var keyParts []string
			emptyKey := true

			for _, idx := range matchIdxs {
				if idx != -1 && idx < len(cols) {
					val := normalizeData(cols[idx])
					keyParts = append(keyParts, val)
					if val != "" {
						emptyKey = false
					}
				} else {
					keyParts = append(keyParts, "")
				}
			}

			if !emptyKey {
				key := strings.Join(keyParts, "|||")
				if val, ok := enrichMap[key]; ok {
					rowVals[len(cols)] = val
				} else {
					rowVals[len(cols)] = "Не найдено"
				}
			} else {
				rowVals[len(cols)] = ""
			}

			cell, _ := excelize.CoordinatesToCellName(1, outRowIdx)
			if err := sw.SetRow(cell, rowVals); err != nil {
				log.Printf("Ошибка записи строки %d: %v", outRowIdx, err)
			}
			outRowIdx++
		}
		rows.Close() // Предотвращаем утечку памяти
		_ = sw.Flush()
	}

	if !foundAnyValidSheet {
		return fmt.Errorf("в целевом файле не найдены колонки для сопоставления")
	}
	return nil
}
