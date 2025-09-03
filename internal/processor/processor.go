package processor

import (
	"fmt"
	"log"
	"strings"
	"time"
	"xlsx-deduper/internal/model"

	"github.com/xuri/excelize/v2"
)

// ReadRecords читает Excel и возвращает слайс model.Record
func ReadRecords(filePath string) ([]model.Record, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, err
	}

	sheet := f.GetSheetName(0)
	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil, err
	}

	var records []model.Record

	for i, row := range rows {
		if i == 0 {
			continue // пропускаем заголовок
		}
		if len(row) < 8 {
			log.Printf("строка %d пропущена: недостаточно колонок", i+1)
			continue
		}

		modified, err := parseDate(row[7])
		if err != nil {
			log.Printf("строка %d ошибка парсинга даты: %v", i+1, err)
			continue
		}

		rec := model.Record{
			Noderef:        row[0],
			LastName:       row[1],
			FirstName:      row[2],
			MiddleName:     row[3],
			Department:     row[4],
			Position:       row[5],
			CounterpartyID: row[6],
			ModifiedSource: modified,
		}
		records = append(records, rec)
	}

	log.Printf("Прочитано записей: %d", len(records))
	return records, nil
}

// parseDate пытается распарсить разные форматы даты
func parseDate(s string) (time.Time, error) {
	// Убираем скобки с названием таймзоны
	if idx := strings.Index(s, "("); idx != -1 {
		s = s[:idx]
	}
	s = strings.TrimSpace(s)

	layout := "Mon Jan 02 2006 15:04:05 MST-0700"
	t, err := time.Parse(layout, s)
	if err != nil {
		return time.Time{}, fmt.Errorf("не удалось распарсить дату: %s", s)
	}
	return t, nil
}

// key формирует ключ для поиска дублей
func key(r model.Record) string {
	return strings.ToLower(r.CounterpartyID) + "|" +
		strings.ToLower(r.LastName) + "|" +
		strings.ToLower(r.FirstName) + "|" +
		strings.ToLower(r.MiddleName) + "|" +
		strings.ToLower(r.Department) + "|" +
		strings.ToLower(r.Position)
}

// ProcessRecords выставляет статус актуальная/архивная
func ProcessRecords(records []model.Record) []model.Record {
	groups := make(map[string][]model.Record)
	for _, r := range records {
		k := key(r)
		groups[k] = append(groups[k], r)
	}

	var result []model.Record

	for _, group := range groups {
		// Сортировка по ModifiedSource (от старого к новому)
		for i := 0; i < len(group)-1; i++ {
			for j := i + 1; j < len(group); j++ {
				if group[i].ModifiedSource.After(group[j].ModifiedSource) {
					group[i], group[j] = group[j], group[i]
				}
			}
		}

		// Первая — архивная, последняя — актуальная
		for i := 0; i < len(group); i++ {
			if i < len(group)-1 {
				group[i].Status = "Архивная"
				group[i].ArchiveNodeRef = group[len(group)-1].Noderef
			} else {
				group[i].Status = "Актуальная"
			}
			result = append(result, group[i])
		}
	}

	return result
}
