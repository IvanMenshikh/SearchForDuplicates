package main

import (
	"log"
	"os"
	"xlsx-deduper/internal/output"
	"xlsx-deduper/internal/processor"
)

func main() {
	const inputFile = "data.xlsx"
	const outputFile = "result.xlsx"

	if _, err := os.Stat(inputFile); os.IsNotExist(err) {
		log.Fatalf("Файл %s не найден в текущей папке", inputFile)
	}

	records, err := processor.ReadRecords(inputFile)
	if err != nil {
		log.Fatalf("Ошибка чтения файла: %v", err)
	}

	processed := processor.ProcessRecords(records)

	err = output.WriteRecords(outputFile, processed)
	if err != nil {
		log.Fatalf("Ошибка записи файла: %v", err)
	}

	log.Println("Обработка завершена. Результат в файле:", outputFile)
}
