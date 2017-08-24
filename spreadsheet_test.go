package gss

import (
	"encoding/json"
	"fmt"
	"net/http"
	"net/url"
	"testing"

	sheets "google.golang.org/api/sheets/v4"

	"github.com/stretchr/testify/assert"
)

func TestNewSpreadsheet(t *testing.T) {
	_, err := NewSpreadsheet(&http.Client{})
	if err != nil {
		t.Error(err)
	}
}

func Test_sheetIdMap(t *testing.T) {
	client, m := newDummyClient(
		map[string]interface{}{
			"sheets": []map[string]interface{}{
				map[string]interface{}{
					"properties": map[string]interface{}{
						"sheetId": 0,
						"title":   "シート1",
					},
				},
				map[string]interface{}{
					"properties": map[string]interface{}{
						"sheetId": 9999,
						"title":   "シート2",
					},
				},
			},
		},
	)
	ss, err := NewSpreadsheet(client)
	if err != nil {
		t.Error(err)
	}
	r, err := ss.sheetIdMap("XXXXXX")
	if err != nil {
		t.Error(err)
	}
	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s", "XXXXXX"), m.req[0].URL.Path)
	assert.Equal(t, url.Values{"alt": []string{"json"}}, m.req[0].URL.Query())
	assert.Equal(t, map[string]int64{
		"シート1": 0,
		"シート2": 9999,
	}, r)
}

func TestSheetCopy(t *testing.T) {
	client, m := newDummyClient(
		map[string]interface{}{
			"sheets": []map[string]interface{}{
				map[string]interface{}{
					"properties": map[string]interface{}{
						"sheetId": 0,
						"title":   "シート1",
					},
				},
				map[string]interface{}{
					"properties": map[string]interface{}{
						"sheetId": 9999,
						"title":   "シート2",
					},
				},
			},
		},
		map[string]interface{}{
			"replies": []interface{}{
				map[string]interface{}{
					"duplicateSheet": map[string]interface{}{
						"properties": map[string]interface{}{
							"sheetId": 9999,
							"title":   "_シート1",
						},
					},
				},
			},
			"spreadsheetId": "XXXXXX",
		},
	)
	ss, err := NewSpreadsheet(client)
	if err != nil {
		t.Error(err)
	}
	err = ss.SheetCopy("XXXXXX", "シート1", "_シート1")
	if err != nil {
		t.Error(err)
	}

	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s", "XXXXXX"), m.req[0].URL.Path)
	assert.Equal(t, url.Values{"alt": []string{"json"}}, m.req[0].URL.Query())

	var reqData interface{}
	err = json.NewDecoder(m.req[1].Body).Decode(&reqData)
	if err != nil {
		t.Error(err)
	}
	expectReqData := map[string]interface{}{
		"requests": []interface{}{
			map[string]interface{}{
				"duplicateSheet": map[string]interface{}{
					"insertSheetIndex": 2.0,
					"newSheetName":     "_シート1",
				},
			},
		},
	}
	assert.Equal(t, expectReqData, reqData)
	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s:batchUpdate", "XXXXXX"), m.req[1].URL.Path)
	assert.Equal(t, url.Values{"alt": []string{"json"}}, m.req[1].URL.Query())
}

func TestSheetDelete(t *testing.T) {
	client, m := newDummyClient(
		map[string]interface{}{
			"sheets": []map[string]interface{}{
				map[string]interface{}{
					"properties": map[string]interface{}{
						"sheetId": 0,
						"title":   "シート1",
					},
				},
				map[string]interface{}{
					"properties": map[string]interface{}{
						"sheetId": 9999,
						"title":   "シート2",
					},
				},
				map[string]interface{}{
					"properties": map[string]interface{}{
						"sheetId": 1234,
						"title":   "_シート1",
					},
				},
			},
		},
		map[string]interface{}{
			"replies":       []interface{}{},
			"spreadsheetId": "XXXXXX",
		},
	)
	ss, err := NewSpreadsheet(client)
	if err != nil {
		t.Error(err)
	}
	err = ss.SheetDelete("XXXXXX", "_シート1")
	if err != nil {
		t.Error(err)
	}

	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s", "XXXXXX"), m.req[0].URL.Path)
	assert.Equal(t, url.Values{"alt": []string{"json"}}, m.req[0].URL.Query())

	var reqData interface{}
	err = json.NewDecoder(m.req[1].Body).Decode(&reqData)
	if err != nil {
		t.Error(err)
	}
	expectReqData := map[string]interface{}{
		"requests": []interface{}{
			map[string]interface{}{
				"deleteSheet": map[string]interface{}{
					"sheetId": 1234.0,
				},
			},
		},
	}
	assert.Equal(t, expectReqData, reqData)
	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s:batchUpdate", "XXXXXX"), m.req[1].URL.Path)
	assert.Equal(t, url.Values{"alt": []string{"json"}}, m.req[1].URL.Query())
}

func TestGetWorksheet(t *testing.T) {
	client, m := newDummyClient(
		map[string]interface{}{
			"range":          "'シート1'!A1:E4",
			"majorDimension": "ROWS",
			"values": []interface{}{
				[]interface{}{"", "column1", "", "column2", "column3"},
				[]interface{}{"", "1", "", "4", "7"},
				[]interface{}{"", "2", "", "5", "8"},
				[]interface{}{"", "3", "", "6", "9"},
			},
		},
	)
	ss, err := NewSpreadsheet(client)
	if err != nil {
		t.Error(err)
	}
	ws, err := ss.GetWorksheet("XXXXXX", "シート1")
	if err != nil {
		t.Error(err)
	}
	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s/values/%s", "XXXXXX", "シート1"), m.req[0].URL.Path)
	assert.Equal(t, url.Values{"alt": []string{"json"}}, m.req[0].URL.Query())
	assert.Equal(t, []map[string]string{
		map[string]string{
			"column1": "1",
			"column2": "4",
			"column3": "7",
		},
		map[string]string{
			"column1": "2",
			"column2": "5",
			"column3": "8",
		},
		map[string]string{
			"column1": "3",
			"column2": "6",
			"column3": "9",
		},
	}, ws.Rows)
}

func TestGetWorksheet_OverRows(t *testing.T) {
	client, m := newDummyClient(
		map[string]interface{}{
			"range":          "'シート1'!A1:F4",
			"majorDimension": "ROWS",
			"values": []interface{}{
				[]interface{}{"", "column1", "", "column2", "column3", ""},
				[]interface{}{"", "1", "", "4", "7", "over1"},
				[]interface{}{"", "2", "", "5", "8", "over2"},
				[]interface{}{"", "3", "", "6", "9", "over3"},
			},
		},
	)
	ss, err := NewSpreadsheet(client)
	if err != nil {
		t.Error(err)
	}
	ws, err := ss.GetWorksheet("XXXXXX", "シート1")
	if err != nil {
		t.Error(err)
	}
	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s/values/%s", "XXXXXX", "シート1"), m.req[0].URL.Path)
	assert.Equal(t, url.Values{"alt": []string{"json"}}, m.req[0].URL.Query())
	assert.Equal(t, []map[string]string{
		map[string]string{
			"column1": "1",
			"column2": "4",
			"column3": "7",
		},
		map[string]string{
			"column1": "2",
			"column2": "5",
			"column3": "8",
		},
		map[string]string{
			"column1": "3",
			"column2": "6",
			"column3": "9",
		},
	}, ws.Rows)
}

func TestWorksheetRefresh(t *testing.T) {
	ws, err := newDummyWorksheet()
	if err != nil {
		t.Error(err)
	}
	client, m := newDummyClient(
		map[string]interface{}{
			"range":          "'シート1'!A1:E4",
			"majorDimension": "ROWS",
			"values": []interface{}{
				[]interface{}{"", "column1", "", "column2", "column3"},
				[]interface{}{"", "1", "", "4", "7"},
				[]interface{}{"", "2", "", "5", "8"},
				[]interface{}{"", "3", "", "6", "9"},
				[]interface{}{"", "11", "", "12", "13"},
			},
		},
	)
	ws.service, err = sheets.New(client)
	if err != nil {
		t.Error(err)
	}
	ws.Refresh()
	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s/values/%s", "XXXXXX", "シート1"), m.req[0].URL.Path)
	assert.Equal(t, url.Values{"alt": []string{"json"}}, m.req[0].URL.Query())
	assert.Equal(t, []map[string]string{
		map[string]string{
			"column1": "1",
			"column2": "4",
			"column3": "7",
		},
		map[string]string{
			"column1": "2",
			"column2": "5",
			"column3": "8",
		},
		map[string]string{
			"column1": "3",
			"column2": "6",
			"column3": "9",
		},
		map[string]string{
			"column1": "11",
			"column2": "12",
			"column3": "13",
		},
	}, ws.Rows)
}

func TestWorksheetAppend(t *testing.T) {
	ws, err := newDummyWorksheet()
	if err != nil {
		t.Error(err)
	}
	client, m := newDummyClient(
		map[string]interface{}{
			"spreadsheetId": "XXXXXX",
			"updates": map[string]interface{}{
				"spreadsheetId":  "XXXXXX",
				"updatedRange":   "'シート1'!A5:E8",
				"updatedRows":    4.000000,
				"updatedColumns": 3.000000,
				"updatedCells":   12.000000,
			},
		},
	)
	ws.service, err = sheets.New(client)
	if err != nil {
		t.Error(err)
	}
	ws.Append([]map[string]string{
		map[string]string{},
		map[string]string{
			"column1": "11",
			"column2": "14",
			"column3": "17",
		},
		map[string]string{
			"column1": "12",
			"column2": "15",
			"column3": "18",
		},
		map[string]string{
			"column1": "13",
			"column2": "16",
			"column3": "19",
		},
	})
	var reqData interface{}
	err = json.NewDecoder(m.req[0].Body).Decode(&reqData)
	if err != nil {
		t.Error(err)
	}
	expectReqData := map[string]interface{}{
		"majorDimension": "ROWS",
		"values": []interface{}{
			[]interface{}{nil, "", nil, "", ""},
			[]interface{}{nil, "11", nil, "14", "17"},
			[]interface{}{nil, "12", nil, "15", "18"},
			[]interface{}{nil, "13", nil, "16", "19"},
		},
	}
	assert.Equal(t, expectReqData, reqData)
	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s/values/%s!A5:append", "XXXXXX", "シート1"), m.req[0].URL.Path)
	assert.Equal(t, url.Values{
		"alt":              []string{"json"},
		"valueInputOption": []string{"USER_ENTERED"},
	}, m.req[0].URL.Query())
	assert.Equal(t, []map[string]string{
		map[string]string{
			"column1": "1",
			"column2": "4",
			"column3": "7",
		},
		map[string]string{
			"column1": "2",
			"column2": "5",
			"column3": "8",
		},
		map[string]string{
			"column1": "3",
			"column2": "6",
			"column3": "9",
		},
		map[string]string{
			"column1": "",
			"column2": "",
			"column3": "",
		},
		map[string]string{
			"column1": "11",
			"column2": "14",
			"column3": "17",
		},
		map[string]string{
			"column1": "12",
			"column2": "15",
			"column3": "18",
		},
		map[string]string{
			"column1": "13",
			"column2": "16",
			"column3": "19",
		},
	}, ws.Rows)
}

func TestWorksheetValues(t *testing.T) {
	ws, err := newDummyWorksheet()
	if err != nil {
		t.Error(err)
	}
	assert.Equal(t, [][]string{
		[]string{"", "1", "", "4", "7"},
		[]string{"", "2", "", "5", "8"},
		[]string{"", "3", "", "6", "9"},
	}, ws.Values())
}

func TestWorksheetDiscardChanges(t *testing.T) {
	ws, err := newDummyWorksheet()
	if err != nil {
		t.Error(err)
	}
	ws.Rows[0]["column1"] = "99"
	ws.Rows[1]["column2"] = "99"
	ws.Rows[2]["column3"] = "99"
	assert.Equal(t, []map[string]string{
		map[string]string{
			"column1": "99",
			"column2": "4",
			"column3": "7",
		},
		map[string]string{
			"column1": "2",
			"column2": "99",
			"column3": "8",
		},
		map[string]string{
			"column1": "3",
			"column2": "6",
			"column3": "99",
		},
	}, ws.Rows)
	ws.DiscardChanges()
	assert.Equal(t, []map[string]string{
		map[string]string{
			"column1": "1",
			"column2": "4",
			"column3": "7",
		},
		map[string]string{
			"column1": "2",
			"column2": "5",
			"column3": "8",
		},
		map[string]string{
			"column1": "3",
			"column2": "6",
			"column3": "9",
		},
	}, ws.Rows)
}

func TestWorksheetUpdate(t *testing.T) {
	ws, err := newDummyWorksheet()
	if err != nil {
		t.Error(err)
	}
	client, m := newDummyClient(
		map[string]interface{}{
			"totalUpdatedColumns": 3.000000,
			"totalUpdatedCells":   3.000000,
			"totalUpdatedSheets":  1.000000,
			"responses": []interface{}{
				map[string]interface{}{
					"spreadsheetId":  "XXXXXX",
					"updatedRange":   "'シート1'!B2",
					"updatedRows":    1.000000,
					"updatedColumns": 1.000000,
					"updatedCells":   1.000000,
				},
				map[string]interface{}{
					"spreadsheetId":  "XXXXXX",
					"updatedRange":   "'シート1'!D3",
					"updatedRows":    1.000000,
					"updatedColumns": 1.000000,
					"updatedCells":   1.000000,
				},
				map[string]interface{}{
					"updatedCells":   1.000000,
					"spreadsheetId":  "XXXXXX",
					"updatedRange":   "'シート1'!E4",
					"updatedRows":    1.000000,
					"updatedColumns": 1.000000,
				},
			},
			"spreadsheetId":    "XXXXXX",
			"totalUpdatedRows": 3.000000,
		},
	)
	ws.service, err = sheets.New(client)
	if err != nil {
		t.Error(err)
	}
	ws.Rows[0]["column1"] = "99"
	ws.Rows[1]["column2"] = "99"
	ws.Rows[2]["column3"] = "99"
	ws.Update()
	var reqData interface{}
	err = json.NewDecoder(m.req[0].Body).Decode(&reqData)
	if err != nil {
		t.Error(err)
	}
	expectReqData := map[string]interface{}{
		"data": []interface{}{
			map[string]interface{}{
				"values": []interface{}{
					[]interface{}{
						"99",
					},
				},
				"majorDimension": "ROWS",
				"range":          "シート1!B2:B2",
			},
			map[string]interface{}{
				"majorDimension": "ROWS",
				"range":          "シート1!D3:D3",
				"values": []interface{}{
					[]interface{}{
						"99",
					},
				},
			},
			map[string]interface{}{
				"values": []interface{}{
					[]interface{}{
						"99",
					},
				},
				"majorDimension": "ROWS",
				"range":          "シート1!E4:E4",
			},
		},
		"valueInputOption": "USER_ENTERED",
	}
	assert.Equal(t, expectReqData, reqData)
	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s/values:batchUpdate", "XXXXXX"), m.req[0].URL.Path)
	assert.Equal(t, url.Values{"alt": []string{"json"}}, m.req[0].URL.Query())
	assert.Equal(t, []map[string]string{
		map[string]string{
			"column1": "99",
			"column2": "4",
			"column3": "7",
		},
		map[string]string{
			"column1": "2",
			"column2": "99",
			"column3": "8",
		},
		map[string]string{
			"column1": "3",
			"column2": "6",
			"column3": "99",
		},
	}, ws.Rows)
	client, m = newDummyClient(
		map[string]interface{}{
			"spreadsheetId": "1zpN3WoM_eRbd0ffpqImm2JkRHKFI6mlfjq9KLHR0aiw",
		},
	)
	ws.service, err = sheets.New(client)
	ws.Update()
	err = json.NewDecoder(m.req[0].Body).Decode(&reqData)
	if err != nil {
		t.Error(err)
	}
	expectReqData = map[string]interface{}{"valueInputOption": "USER_ENTERED"}
	assert.Equal(t, expectReqData, reqData)
	assert.Equal(t, fmt.Sprintf("/v4/spreadsheets/%s/values:batchUpdate", "XXXXXX"), m.req[0].URL.Path)
	assert.Equal(t, url.Values{"alt": []string{"json"}}, m.req[0].URL.Query())
	assert.Equal(t, []map[string]string{
		map[string]string{
			"column1": "99",
			"column2": "4",
			"column3": "7",
		},
		map[string]string{
			"column1": "2",
			"column2": "99",
			"column3": "8",
		},
		map[string]string{
			"column1": "3",
			"column2": "6",
			"column3": "99",
		},
	}, ws.Rows)
}

func TestN2C(t *testing.T) {
	assert.Equal(t, "A", n2c(1))
	assert.Equal(t, "Z", n2c(26))
	assert.Equal(t, "AA", n2c(27))
}
