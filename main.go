package main

import (
	"fmt"
	"io" // Добавлен пакет io для работы с интерфейсами потока
	"log"
	"net/http"
	"os"
	"regexp"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// Регулярное выражение для удаления лишних пробелов внутри текста (например, двойных пробелов в ФИО)
var multipleSpacesRegex = regexp.MustCompile(`\s+`)

func main() {
	port := os.Getenv("PORT")
	if port == "" {
		port = "8080"
	}

	http.HandleFunc("/api/v1/compare", corsMiddleware(compareHandler))
	http.HandleFunc("/api/v1/enrich", corsMiddleware(enrichHandler))

	log.Printf("Сервер запущен на порту %s", port)
	if err := http.ListenAndServe(":"+port, nil); err != nil {
		log.Fatalf("Ошибка запуска сервера: %v", err)
	}
}

// Усиленный corsMiddleware для работы с браузерами
func corsMiddleware(next http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "POST, GET, OPTIONS, PUT, DELETE")
		w.Header().Set("Access-Control-Allow-Headers", "Accept, Content-Type, Content-Length, Accept-Encoding, X-CSRF-Token, Authorization, Origin")
		w.Header().Set("Access-Control-Expose-Headers", "Content-Disposition") // Разрешаем фронтенду видеть имя скачиваемого файла

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}
		next(w, r)
	}
}

// sanitizeKey очищает строку от лишних пробелов по краям и внутри, а также приводит к нижнему регистру
func sanitizeKey(val string) string {
	val = strings.TrimSpace(val)
	val = strings.ToLower(val)
	val = multipleSpacesRegex.ReplaceAllString(val, " ")
	return val
}

// safeOpenExcel принимает любой поток (io.Reader) и безопасно открывает Excel файл с лимитом памяти
func safeOpenExcel(r io.Reader) (*excelize.File, error) {
	opts := excelize.Options{
		UnzipXMLSizeLimit: 1024 * 1024 * 150, // Лимит в 150 MB для защиты от OOM в Render
	}
	return excelize.OpenReader(r, opts)
}

// ==========================================
// ФУНКЦИОНАЛ 1: СВЕРКА ТАБЛИЦ (COMPARE)
// ==========================================
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

	// Убрано жесткое приведение типов .(*os.File)
	f1, err := safeOpenExcel(file1)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 1: "+err.Error(), http.StatusUnprocessableEntity)
		return
	}
	defer f1.Close()

	f2, err := safeOpenExcel(file2)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 2: "+err.Error(), http.StatusUnprocessableEntity)
		return
	}
	defer f2.Close()

	setA, err := buildIDSet(f1, idColumn)
	if err != nil {
		http.Error(w, "Ошибка индексации Файла 1: "+err.Error(), http.StatusBadRequest)
		return
	}

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

	if err := writeDiscrepancies(f1, idColumn, setB, out, sheetName1); err != nil {
		http.Error(w, "Ошибка генерации расхождений Файла 1: "+err.Error(), http.StatusInternalServerError)
		return
	}

	if err := writeDiscrepancies(f2, idColumn, setA, out, sheetName2); err != nil {
		http.Error(w, "Ошибка генерации расхождений Файла 2: "+err.Error(), http.StatusInternalServerError)
		return
	}

	filename := fmt.Sprintf("compare_report_%d.xlsx", time.Now().Unix())
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	if _, err := out.WriteTo(w); err != nil {
		log.Printf("Ошибка отправки отчета клиенту: %v", err)
	}
}

func buildIDSet(f *excelize.File, idColumn string) (map[string]struct{}, error) {
	set := make(map[string]struct{})
	targetIDHeader := sanitizeKey(idColumn)

	for _, sheet := range f.GetSheetList() {
		rows, err := f.Rows(sheet)
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
				for i, col := range cols {
					if sanitizeKey(col) == targetIDHeader {
						idIdx = i
						break
					}
				}
				isFirstRow = false
				continue
			}

			if idIdx != -1 && idIdx < len(cols) {
				val := sanitizeKey(cols[idIdx])
				if val != "" {
					set[val] = struct{}{}
				}
			}
		}
		rows.Close()
	}
	return set, nil
}

func writeDiscrepancies(fIn *excelize.File, idColumn string, targetSet map[string]struct{}, fOut *excelize.File, sheetOut string) error {
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
		isFirstRow := true

		for rows.Next() {
			cols, err := rows.Columns()
			if err != nil {
				break
			}

			if isFirstRow {
				for i, col := range cols {
					if sanitizeKey(col) == targetIDHeader {
						idIdx = i
						break
					}
				}

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
				val := sanitizeKey(cols[idIdx])
				if val != "" {
					if _, exists := targetSet[val]; !exists {
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
		}
		rows.Close()
	}
	return sw.Flush()
}

// ==========================================
// ФУНКЦИОНАЛ 2: ОБОГАЩЕНИЕ ДАННЫХ (ENRICH)
// ==========================================
func enrichHandler(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Метод не поддерживается", http.StatusMethodNotAllowed)
		return
	}

	if err := r.ParseMultipartForm(32 << 20); err != nil {
		http.Error(w, "Ошибка парсинга формы: "+err.Error(), http.StatusBadRequest)
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

	// Убрано жесткое приведение типов .(*os.File)
	f1, err := safeOpenExcel(file1)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 1: "+err.Error(), http.StatusUnprocessableEntity)
		return
	}
	defer f1.Close()

	f2, err := safeOpenExcel(file2)
	if err != nil {
		http.Error(w, "Ошибка чтения Файла 2: "+err.Error(), http.StatusUnprocessableEntity)
		return
	}
	defer f2.Close()

	enrichMap, err := buildEnrichMap(f2, matchCols, sanitizeKey(targetColRaw))
	if err != nil {
		http.Error(w, "Ошибка индексации Базы Данных: "+err.Error(), http.StatusBadRequest)
		return
	}

	out := excelize.NewFile()
	defer out.Close()
	sheetOutName := "Обогащенные данные"
	out.NewSheet(sheetOutName)
	out.DeleteSheet("Sheet1")

	if err := processEnrich(f1, enrichMap, matchCols, targetColRaw, out, sheetOutName); err != nil {
		http.Error(w, "Ошибка обработки данных: "+err.Error(), http.StatusInternalServerError)
		return
	}

	filename := fmt.Sprintf("enriched_report_%d.xlsx", time.Now().Unix())
	w.Header().Set("Content-Disposition", "attachment; filename="+filename)
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	if _, err := out.WriteTo(w); err != nil {
		log.Printf("Ошибка отправки отчета: %v", err)
	}
}

func buildEnrichMap(f *excelize.File, matchCols []string, targetCol string) (map[string]string, error) {
	enrichMap := make(map[string]string)

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
		isFirstRow := true

		for rows.Next() {
			cols, err := rows.Columns()
			if err != nil {
				break
			}

			if isFirstRow {
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
				isFirstRow = false
				continue
			}

			if targetIdx != -1 && targetIdx < len(cols) {
				var keyParts []string
				for _, idx := range matchIdxs {
					if idx != -1 && idx < len(cols) {
						keyParts = append(keyParts, sanitizeKey(cols[idx]))
					} else {
						keyParts = append(keyParts, "")
					}
				}
				key := strings.Join(keyParts, "|||")
				enrichMap[key] = strings.TrimSpace(cols[targetIdx])
			}
		}
		rows.Close()
	}
	return enrichMap, nil
}

func processEnrich(fIn *excelize.File, enrichMap map[string]string, matchCols []string, targetColRaw string, fOut *excelize.File, sheetOut string) error {
	sw, err := fOut.NewStreamWriter(sheetOut)
	if err != nil {
		return err
	}

	outRowIdx := 1
	headerWritten := false

	for _, sheet := range fIn.GetSheetList() {
		rows, err := fIn.Rows(sheet)
		if err != nil {
			continue
		}

		matchIdxs := make([]int, len(matchCols))
		for i := range matchIdxs {
			matchIdxs[i] = -1
		}
		isFirstRow := true

		for rows.Next() {
			cols, err := rows.Columns()
			if err != nil {
				break
			}

			if isFirstRow {
				for i, colName := range cols {
					normCol := sanitizeKey(colName)
					for j, matchCol := range matchCols {
						if normCol == matchCol {
							matchIdxs[j] = i
						}
					}
				}

				if !headerWritten {
					rowVals := make([]interface{}, len(cols)+1)
					for i, v := range cols {
						rowVals[i] = v
					}
					rowVals[len(cols)] = "Найденный " + targetColRaw
					cell, _ := excelize.CoordinatesToCellName(1, outRowIdx)
					_ = sw.SetRow(cell, rowVals)
					outRowIdx++
					headerWritten = true
				}
				isFirstRow = false
				continue
			}

			var keyParts []string
			for _, idx := range matchIdxs {
				if idx != -1 && idx < len(cols) {
					keyParts = append(keyParts, sanitizeKey(cols[idx]))
				} else {
					keyParts = append(keyParts, "")
				}
			}
			key := strings.Join(keyParts, "|||")

			foundID := "Не найдено"
			if val, ok := enrichMap[key]; ok {
				foundID = val
			}

			rowVals := make([]interface{}, len(cols)+1)
			for i, v := range cols {
				rowVals[i] = v
			}
			rowVals[len(cols)] = foundID

			cell, _ := excelize.CoordinatesToCellName(1, outRowIdx)
			_ = sw.SetRow(cell, rowVals)
			outRowIdx++
		}
		rows.Close()
	}

	return sw.Flush()
}