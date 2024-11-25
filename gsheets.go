package gsheets

import (
	"fmt"

	"google.golang.org/api/sheets/v4"
)

func New(service *sheets.Service) *Service {
	return &Service{Service: service}
}

type Service struct {
	Service *sheets.Service
}

func (s *Service) Spreadsheet(id string) (*Spreadsheet, error) {
	spreadsheet, err := s.Service.Spreadsheets.Get(id).Do()
	if err != nil {
		return nil, err
	}
	return &Spreadsheet{
		Service:     s.Service,
		Spreadsheet: spreadsheet,
	}, nil
}

type Spreadsheet struct {
	Service     *sheets.Service
	Spreadsheet *sheets.Spreadsheet
}

func (s *Spreadsheet) Sheet(title string) (*Sheet, error) {
	for _, sheet := range s.Spreadsheet.Sheets {
		if sheet.Properties.Title == title {
			return &Sheet{
				Service:     s.Service,
				Spreadsheet: s.Spreadsheet,
				Sheet:       sheet,
			}, nil
		}
	}
	return nil, fmt.Errorf("sheet not found: %s", title)
}

type Sheet struct {
	Service     *sheets.Service
	Spreadsheet *sheets.Spreadsheet
	Sheet       *sheets.Sheet
}

func (s *Sheet) GetAny(notation string) ([][]any, error) {
	valueRange, err := s.Service.
		Spreadsheets.
		Values.
		Get(
			s.Spreadsheet.SpreadsheetId,
			s.Sheet.Properties.Title+"!"+notation,
		).
		Do()
	if err != nil {
		return nil, err
	}
	return valueRange.Values, nil
}

func (s *Sheet) Get(notation string) ([][]string, error) {
	valuesAny, err := s.GetAny(notation)
	if err != nil {
		return nil, err
	}
	return stringValues(valuesAny), nil
}

func (s *Sheet) UpdateAny(notation string, values [][]any, option string) error {
	_, err := s.Service.
		Spreadsheets.
		Values.
		Update(
			s.Spreadsheet.SpreadsheetId,
			s.Sheet.Properties.Title+"!"+notation,
			&sheets.ValueRange{Values: values},
		).
		ValueInputOption(option).
		Do()
	return err
}

func (s *Sheet) Update(notation string, values [][]string) error {
	return s.UpdateAny(
		notation,
		anyValues(values),
		"USER_ENTERED",
	)
}

func (s *Sheet) AppendAny(notation string, values [][]any, option string) error {
	_, err := s.Service.
		Spreadsheets.
		Values.
		Append(
			s.Spreadsheet.SpreadsheetId,
			s.Sheet.Properties.Title+"!"+notation,
			&sheets.ValueRange{Values: values},
		).
		ValueInputOption(option).
		Do()
	return err
}

func (s *Sheet) Append(notation string, values [][]string) error {
	return s.AppendAny(
		notation,
		anyValues(values),
		"USER_ENTERED",
	)
}

func (s *Sheet) Clear(notation string) error {
	_, err := s.Service.
		Spreadsheets.
		Values.
		Clear(
			s.Spreadsheet.SpreadsheetId,
			s.Sheet.Properties.Title+"!"+notation,
			&sheets.ClearValuesRequest{},
		).
		Do()
	return err
}

func (s *Sheet) InsertRows(start int, rows [][]string) error {
	rowData := make([]*sheets.RowData, len(rows))
	for i, row := range rows {
		cellData := make([]*sheets.CellData, len(row))
		for j, cell := range row {
			cellData[j] = &sheets.CellData{
				UserEnteredValue: &sheets.ExtendedValue{StringValue: &cell},
			}
		}
		rowData[i] = &sheets.RowData{Values: cellData}
	}

	requests := []*sheets.Request{
		{InsertDimension: &sheets.InsertDimensionRequest{
			Range: &sheets.DimensionRange{
				SheetId:    s.Sheet.Properties.SheetId,
				Dimension:  "ROWS",
				StartIndex: int64(start),
				EndIndex:   int64(start + len(rows)),
			},
		}},
		{UpdateCells: &sheets.UpdateCellsRequest{
			Fields: "userEnteredValue",
			Rows:   rowData,
			Start: &sheets.GridCoordinate{
				SheetId:  s.Sheet.Properties.SheetId,
				RowIndex: int64(start),
			},
		}},
	}

	_, err := s.Service.
		Spreadsheets.
		BatchUpdate(
			s.Spreadsheet.SpreadsheetId,
			&sheets.BatchUpdateSpreadsheetRequest{Requests: requests},
		).
		Do()
	return err
}

func stringValues(values [][]any) [][]string {
	result := make([][]string, len(values))
	for i, row := range values {
		result[i] = make([]string, len(row))
		for j, cell := range row {
			result[i][j] = fmt.Sprint(cell)
		}
	}
	return result
}

func anyValues(values [][]string) [][]any {
	result := make([][]any, len(values))
	for i, row := range values {
		result[i] = make([]any, len(row))
		for j, cell := range row {
			result[i][j] = cell
		}
	}
	return result
}
