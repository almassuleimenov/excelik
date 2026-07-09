package main

import (
	"log"
	"net/http"
	"os"
)

func main() {
	port := os.Getenv("PORT")
	if port == "" {
		port = "8080"
	}

	http.HandleFunc("/api/v1/compare", corsMiddleware(compareHandler))
	http.HandleFunc("/api/v1/enrich", corsMiddleware(enrichHandler))

	log.Printf("Сервер запущен на порту %s. Режим: Streaming (SAX) + Memory Safe", port)
	if err := http.ListenAndServe(":"+port, nil); err != nil {
		log.Fatalf("Критическая ошибка запуска сервера: %v", err)
	}
}
