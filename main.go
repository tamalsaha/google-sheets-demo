package main

import (
	"errors"
	"fmt"
	"log"
	"time"

	"golang.org/x/net/context"
	"google.golang.org/api/sheets/v4"
)

type LicenseInfo struct {
	Name      string
	Email     string
	Product   string
	ClusterID string
	Time      string
}

type SheetInfo struct {
	srv            *sheets.Service
	SpreadSheetsID string
	CurrentSheetID int64
}

func main() {
	si := NewSheet("1evwv2ON94R38M-Lkrw8b6dpVSkRYHUWsNOuI7X0_-zA") // Share this sheet with the service account email
	info := LicenseInfo{
		Name:      "Fahim Abrar",
		Email:     "fahimabrar@appscode.com",
		Product:   "Kubeform Community",
		ClusterID: "bad94a42-0210-4c81-b07a-99bae529ec14",
	}

	err := si.insertLicenseInfoInSheet(info)
	if err != nil {
		log.Fatal(err)
	}
}

func NewSheet(spreadsheetId string) *SheetInfo {
	// Set env GOOGLE_APPLICATION_CREDENTIALS to service account json path
	srv, err := sheets.NewService(context.TODO())
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	return &SheetInfo{
		srv:            srv,
		SpreadSheetsID: spreadsheetId,
	}
}

func (si *SheetInfo) getCellData(row, column int64) (string, error) {
	resp, err := si.srv.Spreadsheets.GetByDataFilter(si.SpreadSheetsID, &sheets.GetSpreadsheetByDataFilterRequest{
		IncludeGridData: true,
	}).Do()
	if err != nil {
		return "", fmt.Errorf("unable to retrieve data from sheet: %v", err)
	}

	var val string

	for _, s := range resp.Sheets {
		if s.Properties.SheetId == si.CurrentSheetID {
			val = s.Data[0].RowData[row].Values[column].FormattedValue
		}
	}

	return val, nil
}

// ref: https://developers.google.com/sheets/api/guides/batchupdate
func (si *SheetInfo) updateRowData(row int64, data []string, formatCell bool) error {
	var format *sheets.CellFormat

	if formatCell {
		// for updating header color and making it bold
		format = &sheets.CellFormat{
			TextFormat: &sheets.TextFormat{
				Bold: true,
			},
			BackgroundColor: &sheets.Color{
				Alpha: 1,
				Blue:  149.0 / 255.0,
				Green: 226.0 / 255.0,
				Red:   239.0 / 255.0,
			},
		}
	}

	vals := make([]*sheets.CellData, 0, len(data))
	for i := range data {
		vals = append(vals, &sheets.CellData{
			UserEnteredFormat: format,
			UserEnteredValue: &sheets.ExtendedValue{
				StringValue: &data[i],
			},
		})
	}

	req := []*sheets.Request{
		{
			UpdateCells: &sheets.UpdateCellsRequest{
				Fields: "*",
				Start: &sheets.GridCoordinate{
					ColumnIndex: 0,
					RowIndex:    row,
					SheetId:     si.CurrentSheetID,
				},
				Rows: []*sheets.RowData{
					{
						Values: vals,
					},
				},
			},
		},
	}
	_, err := si.srv.Spreadsheets.BatchUpdate(si.SpreadSheetsID, &sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: false,
		Requests:                     req,
		ResponseIncludeGridData:      false,
	}).Do()
	if err != nil {
		return fmt.Errorf("unable to update: %v", err)
	}

	return nil
}

// ref: https://developers.google.com/sheets/api/guides/batchupdate
func (si *SheetInfo) appendRowData(data []string, formatCell bool) error {
	var format *sheets.CellFormat

	if formatCell {
		// for updating header color and making it bold
		format = &sheets.CellFormat{
			TextFormat: &sheets.TextFormat{
				Bold: true,
			},
			BackgroundColor: &sheets.Color{
				Alpha: 1,
				Blue:  149.0 / 255.0,
				Green: 226.0 / 255.0,
				Red:   239.0 / 255.0,
			},
		}
	}

	vals := make([]*sheets.CellData, 0, len(data))
	for i := range data {
		vals = append(vals, &sheets.CellData{
			UserEnteredFormat: format,
			UserEnteredValue: &sheets.ExtendedValue{
				StringValue: &data[i],
			},
		})
	}

	req := []*sheets.Request{
		{
			AppendCells: &sheets.AppendCellsRequest{
				SheetId: si.CurrentSheetID,
				Fields:  "*",
				Rows: []*sheets.RowData{
					{
						Values: vals,
					},
				},
			},
		},
	}
	_, err := si.srv.Spreadsheets.BatchUpdate(si.SpreadSheetsID, &sheets.BatchUpdateSpreadsheetRequest{
		IncludeSpreadsheetInResponse: false,
		Requests:                     req,
		ResponseIncludeGridData:      false,
	}).Do()
	if err != nil {
		return fmt.Errorf("unable to update: %v", err)
	}

	return nil
}

func (si *SheetInfo) getSheetId(name string) (int64, error) {
	resp, err := si.srv.Spreadsheets.Get(si.SpreadSheetsID).Do()
	if err != nil {
		return -1, fmt.Errorf("unable to retrieve data from sheet: %v", err)
	}
	var id int64
	for _, sheet := range resp.Sheets {
		if sheet.Properties.Title == name {
			id = sheet.Properties.SheetId
		}

	}

	return id, nil
}

func (si *SheetInfo) addNewSheet(name string) error {
	req := sheets.Request{
		AddSheet: &sheets.AddSheetRequest{
			Properties: &sheets.SheetProperties{
				Title: name,
			},
		},
	}

	rbb := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{&req},
	}

	_, err := si.srv.Spreadsheets.BatchUpdate(si.SpreadSheetsID, rbb).Context(context.Background()).Do()
	if err != nil {
		return err
	}

	return nil
}

func (si *SheetInfo) ensureSheet(name string) (int64, error) {
	id, err := si.getSheetId(name)
	if err != nil {
		return 0, err
	}

	if id == 0 {
		err = si.addNewSheet(name)
		if err != nil {
			return 0, err
		}

		id, err = si.getSheetId(name)
		if err != nil {
			return 0, err
		}

		si.CurrentSheetID = id

		err = si.ensureHeader()
		if err != nil {
			return 0, err
		}

		return id, nil
	}

	si.CurrentSheetID = id
	return id, nil
}

func (si *SheetInfo) ensureHeader() error {
	headers := []string{"SL", "Name", "Email", "ClusterID", "Time"}
	return si.updateRowData(0, headers, true)
}

func (si *SheetInfo) findEmptyCell() (int64, error) {
	resp, err := si.srv.Spreadsheets.GetByDataFilter(si.SpreadSheetsID, &sheets.GetSpreadsheetByDataFilterRequest{
		IncludeGridData: true,
	}).Do()
	if err != nil {
		return 0, fmt.Errorf("unable to retrieve data from sheet: %v", err)
	}

	for _, s := range resp.Sheets {
		if s.Properties.SheetId == si.CurrentSheetID {
			return int64(len(s.Data[0].RowData)), nil
		}
	}

	return 0, errors.New("no empty cell found")
}

func (si *SheetInfo) insertLicenseInfoInSheet(info LicenseInfo) error {
	_, err := si.ensureSheet(info.Product)
	if err != nil {
		log.Fatal(err)
	}

	vals := []string{
		"1",
		info.Name,
		info.Email,
		info.ClusterID,
		time.Now().UTC().Format(time.RFC3339),
	}
	return si.appendRowData(vals, false)
}
