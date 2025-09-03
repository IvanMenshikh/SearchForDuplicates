package output

import (
	"fmt"
	"log"
	"xlsx-deduper/internal/model"

	"github.com/xuri/excelize/v2"
)

func WriteRecords(filePath string, records []model.Record) error {
	f := excelize.NewFile()
	sheet := f.GetSheetName(0)

	headers := []string{
		"Noderef", "Фамилия", "Имя", "Отчество",
		"Департамент", "Должность", "ID контрагента",
		"modified-source UTC0", "Статус", "ArchiveNodeRef",
	}

	headerStyle, _ := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"#FFC000"},
			Pattern: 1,
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 2}, // сдвоенная нижняя
		},
	})

	cellBorderStyle, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	for i, h := range headers {
		col, _ := excelize.ColumnNumberToName(i + 1)
		cell := fmt.Sprintf("%s1", col)
		f.SetCellValue(sheet, cell, h)
		f.SetCellStyle(sheet, cell, cell, headerStyle)
	}

	for i, rec := range records {
		row := i + 2
		values := []interface{}{
			rec.Noderef,
			rec.LastName,
			rec.FirstName,
			rec.MiddleName,
			rec.Department,
			rec.Position,
			rec.CounterpartyID,
			rec.ModifiedSource.Format("2006-01-02 15:04:05"),
			rec.Status,
			rec.ArchiveNodeRef,
		}

		for j, val := range values {
			col, _ := excelize.ColumnNumberToName(j + 1)
			cell := fmt.Sprintf("%s%d", col, row)
			f.SetCellValue(sheet, cell, val)
			f.SetCellStyle(sheet, cell, cell, cellBorderStyle)
		}
	}

	err := f.SetPanes(sheet, &excelize.Panes{
		Freeze:      true,
		XSplit:      0,
		YSplit:      1,
		TopLeftCell: "A2",
		ActivePane:  "bottomLeft",
	})
	if err != nil {
		log.Fatalf("Ошибка заморозки панели: %v", err)
	}

	return f.SaveAs(filePath)
}
