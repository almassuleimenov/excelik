package main

import (
	"fmt"
	"log"
	"net/http"
	"os"

	"github.com/xuri/excelize/v2"
)

const maxUploadSize = 50 << 20 // Хард-лимит 50 MB на весь запрос

func compareHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Метод не поддерживается", http.StatusMethodNotAllowed)
		return
	}

	r.Body = http.MaxBytesReader(w, r.Body, maxUploadSize)
	if err := r.ParseMultipartForm(15 << 20); err != nil {
		http.Error(w, "Ошибка формы (возможно превышен лимит 50МБ): "+err.Error(), http.StatusBadRequest)
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

	tempPath1, err := saveToTempFile(file1, "file1")
	if err != nil {
		http.Error(w, "Ошибка сохранения Файла 1: "+err.Error(), http.StatusInternalServerError)
		return
	}
	defer os.Remove(tempPath1)

	tempPath2, err := saveToTempFile(file2, "file2")
	if err != nil {
		http.Error(w, "Ошибка сохранения Файла 2: "+err.Error(), http.StatusInternalServerError)
		return
	}
	defer os.Remove(tempPath2)

	f1, err := excelize.OpenFile(tempPath1)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 1: "+err.Error(), http.StatusUnprocessableEntity)
		return
	}
	defer f1.Close()

	f2, err := excelize.OpenFile(tempPath2)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 2: "+err.Error(), http.StatusUnprocessableEntity)
		return
	}
	defer f2.Close()

	setA, err := buildIDSetStream(f1, idColumn)
	if err != nil {
		http.Error(w, "Ошибка индексации Файла 1: "+err.Error(), http.StatusBadRequest)
		return
	}

	setB, err := buildIDSetStream(f2, idColumn)
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

	if err := writeDiscrepanciesStream(f1, idColumn, setB, out, sheetName1); err != nil {
		http.Error(w, "Ошибка генерации расхождений Файла 1: "+err.Error(), http.StatusInternalServerError)
		return
	}

	if err := writeDiscrepanciesStream(f2, idColumn, setA, out, sheetName2); err != nil {
		http.Error(w, "Ошибка генерации расхождений Файла 2: "+err.Error(), http.StatusInternalServerError)
		return
	}

	tempOutPath, err := createSafeTempOutFile("compare_report")
	if err != nil {
		http.Error(w, "Ошибка создания отчета", http.StatusInternalServerError)
		return
	}
	defer os.Remove(tempOutPath)

	if err := out.SaveAs(tempOutPath); err != nil {
		log.Printf("Ошибка сохранения итогового файла: %v", err)
		http.Error(w, "Ошибка при компиляции отчета", http.StatusInternalServerError)
		return
	}

	w.Header().Set("Content-Disposition", "attachment; filename=compare_report.xlsx")
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	http.ServeFile(w, r, tempOutPath)
}

func buildIDSetStream(f *excelize.File, idColumn string) (map[string]struct{}, error) {
	set := make(map[string]struct{})
	targetIDHeader := sanitizeKey(idColumn)
	foundValidSheet := false

	for _, sheet := range f.GetSheetList() {
		rows, err := f.Rows(sheet)
		if err != nil {
			continue
		}

		idIdx := -1
		headerFound := false

		for rows.Next() {
			cols, err := rows.Columns()
			if err != nil {
				continue
			}

			if !headerFound {
				for i, col := range cols {
					if sanitizeKey(col) == targetIDHeader {
						idIdx = i
						headerFound = true
						foundValidSheet = true
						break
					}
				}
				continue
			}

			if idIdx != -1 && idIdx < len(cols) {
				val := sanitizeKey(cols[idIdx])
				if val != "" {
					set[val] = struct{}{}
				}
			}
		}
		rows.Close() // КРИТИЧЕСКИ ВАЖНО: очистка ресурсов SAX парсера
	}

	if !foundValidSheet {
		return nil, fmt.Errorf("колонка '%s' не найдена ни на одном листе", idColumn)
	}
	return set, nil
}

func writeDiscrepanciesStream(fIn *excelize.File, idColumn string, targetSet map[string]struct{}, fOut *excelize.File, sheetOut string) error {
	sw, err := fOut.NewStreamWriter(sheetOut)
	if err != nil {
		return err
	}

	outRowIdx := 1
	headerWritten := false
	targetIDHeader := sanitizeKey(idColumn)

	for _, sheetIn := range fIn.GetSheetList() {
		rows, err := fIn.Rows(sheetIn)
		if err != nil {
			continue
		}

		idIdx := -1
		headerFound := false

		for rows.Next() {
			cols, err := rows.Columns()
			if err != nil {
				continue
			}

			if !headerFound {
				for i, col := range cols {
					if sanitizeKey(col) == targetIDHeader {
						idIdx = i
						headerFound = true
						break
					}
				}

				if headerFound && !headerWritten {
					rowVals := make([]interface{}, len(cols))
					for i, col := range cols {
						rowVals[i] = col
					}
					cell, _ := excelize.CoordinatesToCellName(1, outRowIdx)
					if err := sw.SetRow(cell, rowVals); err != nil {
						log.Printf("Ошибка записи заголовка: %v", err)
					}
					outRowIdx++
					headerWritten = true
				}
				continue
			}

			if idIdx != -1 && idIdx < len(cols) {
				val := sanitizeKey(cols[idIdx])
				if val != "" {
					if _, exists := targetSet[val]; !exists {
						rowVals := make([]interface{}, len(cols))
						for i, col := range cols {
							rowVals[i] = col
						}
						cell, _ := excelize.CoordinatesToCellName(1, outRowIdx)
						if err := sw.SetRow(cell, rowVals); err != nil {
							log.Printf("Ошибка записи строки %d: %v", outRowIdx, err)
						}
						outRowIdx++
					}
				}
			}
		}
		rows.Close() // Очистка ресурсов после прохода по листу
	}
	return sw.Flush()
}
