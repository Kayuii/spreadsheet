package spreadsheet

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"net/url"
	"os"
	"strings"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
)

const (
	baseURL = "https://sheets.googleapis.com/v4"

	// Scope is the API scope for viewing and managing your Google Spreadsheet data.
	// Useful for generating JWT values.
	Scope = "https://spreadsheets.google.com/feeds"

	// SecretFileName is used to get client.
	SecretFileName = "client_secret.json"

	// DriveScope See, edit, create, and delete all of your Google Drive files
	DriveScope = "https://www.googleapis.com/auth/drive"

	// DriveFileScope View and manage Google Drive files and folders that you have opened
	// or created with this app
	DriveFileScope = "https://www.googleapis.com/auth/drive.file"

	// DriveReadonlyScope See and download all your Google Drive files
	DriveReadonlyScope = "https://www.googleapis.com/auth/drive.readonly"

	// SpreadsheetsScope See, edit, create, and delete your spreadsheets in Google Drive
	SpreadsheetsScope = "https://www.googleapis.com/auth/spreadsheets"

	// SpreadsheetsReadonlyScope View your Google Spreadsheets
	SpreadsheetsReadonlyScope = "https://www.googleapis.com/auth/spreadsheets.readonly"
)

// NewServiceForCLI returns a gsheets client.
// This function is intended for CLI tools.
func NewServiceForCLI(ctx context.Context, authFile string) (s *Service, err error) {

	cb, err := ioutil.ReadFile(authFile)
	if err != nil {
		return nil, fmt.Errorf("Unable to read client secret file: %v", err)
	}

	config, err := google.ConfigFromJSON(cb, SpreadsheetsScope)
	if err != nil {
		return nil, fmt.Errorf("Unable to parse client secret file to config: %v", err)
	}

	tokenFile := "token.json"
	tb, err := ioutil.ReadFile(tokenFile)

	var token string
	if err == nil {
		token = string(tb)
	} else {
		// if there are no token file, get from Web

		authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
		fmt.Printf("Go to the following link in your browser then type the "+
			"authorization code: \n%v\n", authURL)

		var authCode string
		if _, err := fmt.Scan(&authCode); err != nil {
			return nil, fmt.Errorf("Unable to read authorization code: %v", err)
		}

		tok, err := config.Exchange(oauth2.NoContext, authCode)
		if err != nil {
			return nil, fmt.Errorf("Unable to retrieve token from web: %v", err)
		}

		b := &bytes.Buffer{}
		json.NewEncoder(b).Encode(tok)
		token = b.String()

		// save token
		fmt.Printf("Saving credential file to: %s\n", tokenFile)
		f, err := os.OpenFile(tokenFile, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
		defer f.Close()
		if err != nil {
			return nil, fmt.Errorf("Unable to cache oauth token: %v", err)
		}
		fmt.Fprint(f, token)
	}

	tok := &oauth2.Token{}
	if err := json.NewDecoder(strings.NewReader(token)).Decode(tok); err != nil {
		return nil, fmt.Errorf("Unable to parse json to token: %v", err)
	}
	s = NewServiceWithClient(config.Client(ctx, tok))
	return
}

// NewService makes a new service with the secret file.
func NewService() (s *Service, err error) {
	data, err := ioutil.ReadFile(SecretFileName)
	if err != nil {
		return
	}

	conf, err := google.JWTConfigFromJSON(data, Scope)
	if err != nil {
		return
	}

	s = NewServiceWithClient(conf.Client(oauth2.NoContext))
	return
}

// NewServiceWithClient makes a new service by the client.
func NewServiceWithClient(client *http.Client) *Service {
	return &Service{
		baseURL: baseURL,
		client:  client,
	}
}

// Service represents a Sheets API service instance.
// Service is the main entry point into using this package.
type Service struct {
	baseURL string
	client  *http.Client
}

// CreateSpreadsheet creates a spreadsheet with the given title
func (s *Service) CreateSpreadsheet(spreadsheet Spreadsheet) (resp Spreadsheet, err error) {
	sheets := make([]map[string]interface{}, 1)
	for s := range spreadsheet.Sheets {
		sheet := spreadsheet.Sheets[s]
		sheets = append(sheets, map[string]interface{}{"properties": map[string]interface{}{"title": sheet.Properties.Title}})
	}
	body, err := s.post("/spreadsheets", map[string]interface{}{
		"properties": map[string]interface{}{
			"title": spreadsheet.Properties.Title,
		},
		"sheets": sheets,
	})
	if err != nil {
		return
	}
	err = json.Unmarshal([]byte(body), &resp)
	if err != nil {
		return
	}
	return s.FetchSpreadsheet(resp.ID)
}

// FetchSpreadsheet fetches the spreadsheet by the id.
func (s *Service) FetchSpreadsheet(id string) (spreadsheet Spreadsheet, err error) {
	fields := "spreadsheetId,properties.title,sheets(properties,data.rowData.values(userEnteredValue))"
	fields = url.QueryEscape(fields)
	path := fmt.Sprintf("/spreadsheets/%s?fields=%s", id, fields)
	body, err := s.get(path)
	if err != nil {
		return
	}
	err = json.Unmarshal(body, &spreadsheet)
	if err != nil {
		return
	}
	spreadsheet.service = s
	return
}

// ReloadSpreadsheet reloads the spreadsheet
func (s *Service) ReloadSpreadsheet(spreadsheet *Spreadsheet) (err error) {
	newSpreadsheet, err := s.FetchSpreadsheet(spreadsheet.ID)
	if err != nil {
		return
	}
	spreadsheet.Properties = newSpreadsheet.Properties
	spreadsheet.Sheets = newSpreadsheet.Sheets
	return
}

// UpdateSpreadsheetTitle update spreadsheet title
func (s *Service) UpdateSpreadsheetTitle(spreadsheet *Spreadsheet, properties Properties) (err error) {
	r, err := newUpdateRequest(spreadsheet)
	if err != nil {
		return
	}
	err = r.UpdateSpreadsheetProperties(&properties).Do()
	if err != nil {
		return
	}
	err = s.ReloadSpreadsheet(spreadsheet)
	return
}

// UpdateSheetTitle update spreadsheet title
func (s *Service) UpdateSheetTitle(sheet *Sheet, sheetProperties SheetProperties) (err error) {
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.UpdateSheetProperties(sheet, &sheetProperties).Do()
	if err != nil {
		return
	}
	err = s.ReloadSpreadsheet(sheet.Spreadsheet)
	return
}

// AddSheet adds a sheet
func (s *Service) AddSheet(spreadsheet *Spreadsheet, sheetProperties SheetProperties) (err error) {
	r, err := newUpdateRequest(spreadsheet)
	if err != nil {
		return
	}
	err = r.AddSheet(sheetProperties).Do()
	if err != nil {
		return
	}
	err = s.ReloadSpreadsheet(spreadsheet)
	return
}

// DeleteSheet deletes the sheet
func (s *Service) DeleteSheet(spreadsheet *Spreadsheet, sheetID uint) (err error) {
	r, err := newUpdateRequest(spreadsheet)
	if err != nil {
		return
	}
	err = r.DeleteSheet(sheetID).Do()
	if err != nil {
		return
	}
	err = s.ReloadSpreadsheet(spreadsheet)
	return
}

// SyncSheet updates sheet
func (s *Service) SyncSheet(sheet *Sheet) (err error) {
	if sheet.newMaxRow > sheet.Properties.GridProperties.RowCount ||
		sheet.newMaxColumn > sheet.Properties.GridProperties.ColumnCount {
		err = s.ExpandSheet(sheet, sheet.newMaxRow, sheet.newMaxColumn)
		if err != nil {
			return
		}
	}
	err = s.syncCells(sheet)
	if err != nil {
		return
	}
	sheet.modifiedCells = []*Cell{}
	sheet.newMaxRow = sheet.Properties.GridProperties.RowCount
	sheet.newMaxColumn = sheet.Properties.GridProperties.ColumnCount
	return
}

// ExpandSheet expands the range of the sheet
func (s *Service) ExpandSheet(sheet *Sheet, row, column uint) (err error) {
	props := sheet.Properties
	props.GridProperties.RowCount = row
	props.GridProperties.ColumnCount = column

	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.UpdateSheetProperties(sheet, &props).Do()
	if err != nil {
		return
	}
	sheet.newMaxRow = row
	sheet.newMaxColumn = column
	return
}

// AppendCells inserts rows into the sheet
func (s *Service) AppendCells(sheet *Sheet, rows [][]Cell) (err error) {
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.AppendCells(sheet, rows).Do()
	return
}

// InsertRows inserts rows into the sheet
func (s *Service) InsertRows(sheet *Sheet, start, end int) (err error) {
	sheet.Properties.GridProperties.RowCount -= uint(end - start)
	sheet.newMaxRow -= uint(end - start)
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.InsertDimension(sheet, "ROWS", start, end).Do()
	return
}

// DeleteRows deletes rows from the sheet
func (s *Service) DeleteRows(sheet *Sheet, start, end int) (err error) {
	sheet.Properties.GridProperties.RowCount -= uint(end - start)
	sheet.newMaxRow -= uint(end - start)
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.DeleteDimension(sheet, "ROWS", start, end).Do()
	return
}

// DeleteColumns deletes columns from the sheet
func (s *Service) DeleteColumns(sheet *Sheet, start, end int) (err error) {
	sheet.Properties.GridProperties.ColumnCount -= uint(end - start)
	sheet.newMaxRow -= uint(end - start)
	r, err := newUpdateRequest(sheet.Spreadsheet)
	if err != nil {
		return
	}
	err = r.DeleteDimension(sheet, "COLUMNS", start, end).Do()
	return
}

func (s *Service) syncCells(sheet *Sheet) (err error) {
	path := fmt.Sprintf("/spreadsheets/%s/values:batchUpdate", sheet.Spreadsheet.ID)
	params := map[string]interface{}{
		"valueInputOption": "USER_ENTERED",
		"data":             make([]map[string]interface{}, 0, len(sheet.modifiedCells)),
	}
	for _, cell := range sheet.modifiedCells {
		valueRange := map[string]interface{}{
			"range":          sheet.Properties.Title + "!" + cell.Pos(),
			"majorDimension": "COLUMNS",
			"values": [][]string{
				[]string{
					cell.Value,
				},
			},
		}
		params["data"] = append(params["data"].([]map[string]interface{}), valueRange)
	}
	_, err = sheet.Spreadsheet.service.post(path, params)
	return
}

func (s *Service) get(path string) (body []byte, err error) {
	resp, err := s.client.Get(baseURL + path)
	if err != nil {
		return
	}
	body, err = ioutil.ReadAll(resp.Body)
	resp.Body.Close()
	if err != nil {
		return
	}
	err = s.checkError(body)
	return
}

func (s *Service) post(path string, params map[string]interface{}) (body string, err error) {
	reqBody, err := json.Marshal(params)
	if err != nil {
		return
	}
	resp, err := s.client.Post(baseURL+path, "application/json", bytes.NewReader(reqBody))
	if err != nil {
		return
	}
	bytes, err := ioutil.ReadAll(resp.Body)
	resp.Body.Close()
	if err != nil {
		return
	}
	err = s.checkError(bytes)
	if err != nil {
		return
	}
	body = string(bytes)
	return
}

func (s *Service) checkError(body []byte) (err error) {
	var res map[string]interface{}
	err = json.Unmarshal(body, &res)
	if err != nil {
		return
	}
	resErr, hasErr := res["error"].(map[string]interface{})
	if !hasErr {
		return
	}
	code := resErr["code"].(float64)
	message := resErr["message"].(string)
	status := resErr["status"].(string)
	err = fmt.Errorf("error status: %s, code:%d, message: %s", status, int(code), message)
	return
}
