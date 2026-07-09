package main

import (
	"fmt"
	"io"
	"net/http"
	"os"
	"strings"
)

func corsMiddleware(next http.HandlerFunc) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "POST, GET, OPTIONS, PUT, DELETE")
		w.Header().Set("Access-Control-Allow-Headers", "Accept, Content-Type, Content-Length, Accept-Encoding, X-CSRF-Token, Authorization, Origin")
		w.Header().Set("Access-Control-Expose-Headers", "Content-Disposition")

		if r.Method == "OPTIONS" {
			w.WriteHeader(http.StatusOK)
			return
		}
		next(w, r)
	}
}

func sanitizeKey(val string) string {
	val = strings.ToLower(val)
	return strings.Join(strings.Fields(val), " ")
}

func normalizeData(val string) string {
	val = strings.ToLower(val)
	val = strings.ReplaceAll(val, " ", "")
	val = strings.ReplaceAll(val, "\u00A0", "")
	val = strings.ReplaceAll(val, "\t", "")
	return val
}

func saveToTempFile(r io.Reader, prefix string) (string, error) {
	tempFile, err := os.CreateTemp("", fmt.Sprintf("%s-*.xlsx", prefix))
	if err != nil {
		return "", err
	}
	defer tempFile.Close()

	if _, err := io.Copy(tempFile, r); err != nil {
		return "", err
	}
	return tempFile.Name(), nil
}

// createSafeTempOutFile создает безопасный временный файл для итогового отчета
func createSafeTempOutFile(prefix string) (string, error) {
	tempOut, err := os.CreateTemp("", fmt.Sprintf("%s-*.xlsx", prefix))
	if err != nil {
		return "", err
	}
	path := tempOut.Name()
	tempOut.Close() // Закрываем дескриптор, excelize сам откроет его при SaveAs
	return path, nil
}
