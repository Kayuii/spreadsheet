package spreadsheet

import (
	"errors"
	"fmt"
	"log"
	"strings"
)

func newUpdateRequest(spreadsheet *Spreadsheet) (r *updateRequest, err error) {
	if spreadsheet == nil {
		err = errors.New("spreadsheet must not be nil")
		return
	}
	r = &updateRequest{
		spreadsheet: spreadsheet,
		body: map[string][]map[string]interface{}{
			"requests": make([]map[string]interface{}, 0, 1),
		},
	}
	return
}

type updateRequest struct {
	spreadsheet *Spreadsheet
	body        map[string][]map[string]interface{}
}

func (r *updateRequest) Do() (err error) {
	if len(r.body["requests"]) == 0 {
		err = errors.New("Requests must not be empty")
		return
	}
	path := fmt.Sprintf("/spreadsheets/%s:batchUpdate", r.spreadsheet.ID)
	params := make(map[string]interface{}, len(r.body))
	for k, v := range r.body {
		params[k] = v
	}
	_, err = r.spreadsheet.service.post(path, params)
	return
}

func (r *updateRequest) UpdateSpreadsheetProperties(spreadsheetProperties *Properties) (ret *updateRequest) {
	ret = r
	params := map[string]interface{}{}
	fields := []string{}
	if spreadsheetProperties.Title != "" {
		params["title"] = spreadsheetProperties.Title
		fields = append(fields, "title")
	}
	if spreadsheetProperties.Locale != "" {
		params["locale"] = spreadsheetProperties.Locale
		fields = append(fields, "locale")
	}
	if spreadsheetProperties.AutoRecalc != "" {
		params["autoRecalc"] = spreadsheetProperties.Locale
		fields = append(fields, "autoRecalc")
	}
	if spreadsheetProperties.TimeZone != "" {
		params["timezone"] = spreadsheetProperties.Locale
		fields = append(fields, "timezone")
	}
	if len(fields) == 0 {
		return
	}
	r.body["requests"] = append(r.body["requests"], map[string]interface{}{
		"updateSpreadsheetProperties": map[string]interface{}{
			"properties": params,
			"fields":     strings.Join(fields, ","),
		},
	})
	return
}

func (r *updateRequest) UpdateSheetProperties(sheet *Sheet, sheetProperties *SheetProperties) (ret *updateRequest) {
	ret = r
	params := map[string]interface{}{
		"sheetId": sheet.Properties.ID,
	}
	fields := []string{}
	if sheetProperties.Title != sheet.Properties.Title {
		params["title"] = sheetProperties.Title
		fields = append(fields, "title")
	}
	if sheetProperties.Index != sheet.Properties.Index {
		params["index"] = sheetProperties.Index
		fields = append(fields, "index")
	}
	gridParams := make(map[string]interface{}, 0)
	props := sheetProperties.GridProperties
	currentProps := sheet.Properties.GridProperties
	if props.RowCount != currentProps.RowCount {
		gridParams["rowCount"] = props.RowCount
		fields = append(fields, "gridProperties.rowCount")
	}
	if props.ColumnCount != currentProps.ColumnCount {
		gridParams["columnCount"] = props.ColumnCount
		fields = append(fields, "gridProperties.columnCount")
	}
	if props.FrozenRowCount != currentProps.FrozenRowCount {
		gridParams["frozenRowCount"] = props.FrozenRowCount
		fields = append(fields, "gridProperties.frozenRowCount")
	}
	if props.FrozenColumnCount != currentProps.FrozenColumnCount {
		gridParams["frozenColumnCount"] = props.FrozenColumnCount
		fields = append(fields, "gridProperties.frozenColumnCount")
	}
	if props.HideGridlines != currentProps.HideGridlines {
		gridParams["hideGridlines"] = props.HideGridlines
		fields = append(fields, "gridProperties.hideGridlines")
	}
	if len(gridParams) > 0 {
		params["gridProperties"] = gridParams
	}
	if sheetProperties.Hidden != sheet.Properties.Hidden {
		params["hidden"] = sheetProperties.Hidden
		fields = append(fields, "hidden")
	}
	if sheetProperties.TabColor != sheet.Properties.TabColor {
		params["tabColor"] = sheetProperties.TabColor
		fields = append(fields, "tabColor")
	}
	if sheetProperties.RightToLeft != sheet.Properties.RightToLeft {
		params["rightToLeft"] = sheet.Properties.RightToLeft
		fields = append(fields, "rightToLeft")
	}
	if len(fields) == 0 {
		return
	}
	r.body["requests"] = append(r.body["requests"], map[string]interface{}{
		"updateSheetProperties": map[string]interface{}{
			"properties": params,
			"fields":     strings.Join(fields, ","),
		},
	})
	return
}

func (r *updateRequest) UpdateDimensionProperties() {

}

func (r *updateRequest) UpdateNamedRange() {

}

func (r *updateRequest) RepeatCell() {

}

func (r *updateRequest) AddNamedRange() {

}

func (r *updateRequest) DeleteNamedRange() {

}

func (r *updateRequest) AddSheet(sheetProperties SheetProperties) *updateRequest {
	r.body["requests"] = append(r.body["requests"], map[string]interface{}{
		"addSheet": map[string]interface{}{
			"properties": sheetProperties,
		},
	})
	return r
}

func (r *updateRequest) DeleteSheet(sheetID uint) *updateRequest {
	r.body["requests"] = append(r.body["requests"], map[string]interface{}{
		"deleteSheet": map[string]interface{}{
			"sheetId": sheetID,
		},
	})
	return r
}

func (r *updateRequest) AutoFill() {

}

func (r *updateRequest) CutPaste() {

}

func (r *updateRequest) CopyPaste() {

}

func (r *updateRequest) MergeCells() {

}

func (r *updateRequest) UnmergeCells() {

}

func (r *updateRequest) UpdateBorders() {

}

func (r *updateRequest) UpdateCells() {

}

func (r *updateRequest) AddFilterView() {

}

func (r *updateRequest) AppendCells(sheet *Sheet, rows [][]Cell) *updateRequest {

	// _rows := make([]interface{}, len(rows))
	for k, v := range rows {

		log.Printf("kv key: %v  value: %v", k, v)

		// _rows = append(_rows, )
		// record := make()
		for a, b := range v {
			log.Printf("ab key: %v  value: %v", a, b)
			// record = append(record, b.Value)
		}

	}

	r.body["requests"] = append(r.body["requests"], map[string]interface{}{
		"appendCells": map[string]interface{}{
			"sheetId": sheet.Properties.ID,
			// "rows":    rows,
			"fields": "*", //strings.Join(fields, ","),
		},
	})
	return r
}

func (r *updateRequest) ClearBasicFilter() {

}

// DeleteDemension deletes rows or columns
func (r *updateRequest) DeleteDimension(sheet *Sheet, dimension string, start, end int) (ret *updateRequest) {
	r.body["requests"] = append(r.body["requests"], map[string]interface{}{
		"deleteDimension": map[string]interface{}{
			"range": map[string]interface{}{
				"sheetId":    sheet.Properties.ID,
				"dimension":  dimension,
				"startIndex": start,
				"endIndex":   end,
			},
		},
	})
	return r
}

func (r *updateRequest) DeleteEmbeddedObject() {

}

func (r *updateRequest) DeleteFilterView() {

}

func (r *updateRequest) DuplicateFilterView() {

}

func (r *updateRequest) DuplicateSheet() {

}

func (r *updateRequest) FindReplace() {

}

func (r *updateRequest) InsertDimension(sheet *Sheet, dimension string, start, end int) (ret *updateRequest) {
	r.body["requests"] = append(r.body["requests"], map[string]interface{}{
		"insertDimension": map[string]interface{}{
			"range": map[string]interface{}{
				"sheetId":    sheet.Properties.ID,
				"dimension":  dimension,
				"startIndex": start,
				"endIndex":   end,
			},
			// "inheritFromBefore": false
		},
	})
	return r
}

func (r *updateRequest) MoveDimension() {

}

func (r *updateRequest) UpdateEmbeddedObjectPosition() {

}

func (r *updateRequest) PasteData() {

}

func (r *updateRequest) TextToColumns() {

}

func (r *updateRequest) UpdateFilterView() {

}

func (r *updateRequest) AppendDimension() {

}

func (r *updateRequest) AddConditionalFormatRule() {

}

func (r *updateRequest) UpdateConditionalFormatRule() {

}

func (r *updateRequest) DeleteConditionalFormatRule() {

}

func (r *updateRequest) SortRange() {

}

func (r *updateRequest) SetDataValidation() {

}

func (r *updateRequest) SetBasicFilter() {

}

func (r *updateRequest) AddProtectedRange() {

}

func (r *updateRequest) UpdateProtectedRange() {

}

func (r *updateRequest) DeleteProtectedRange() {

}

func (r *updateRequest) AutoResizeDimensions() {

}

func (r *updateRequest) AddChart() {

}

func (r *updateRequest) UpdateChartSpec() {

}

func (r *updateRequest) UpdateBanding() {

}

func (r *updateRequest) AddBanding() {

}

func (r *updateRequest) DeleteBanding() {

}
