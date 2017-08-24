package gss

import (
	"fmt"
	"net/http"

	sheets "google.golang.org/api/sheets/v4"
)

type Spreadsheet struct {
	service *sheets.Service
}

func NewSpreadsheet(client *http.Client) (*Spreadsheet, error) {
	service, err := sheets.New(client)
	if err != nil {
		return nil, err
	}
	return &Spreadsheet{service: service}, nil
}

func (ss *Spreadsheet) GetWorksheet(key, sheetName string) (*Worksheet, error) {
	ws := &Worksheet{
		service:          ss.service,
		sheetKey:         key,
		sheetName:        sheetName,
		MajorDimension:   "ROWS",
		ValueInputOption: "USER_ENTERED",
	}
	if err := ws.Refresh(); err != nil {
		return nil, err
	}
	return ws, nil
}

func (ss *Spreadsheet) sheetIdMap(key string) (map[string]int64, error) {
	r, err := ss.service.Spreadsheets.Get(key).Do()
	if err != nil {
		return nil, err
	}
	var res = make(map[string]int64)
	for _, s := range r.Sheets {
		res[s.Properties.Title] = s.Properties.SheetId
	}
	return res, nil
}

func (ss *Spreadsheet) SheetCopy(key, srcName, dstName string) error {
	sheetIdMap, err := ss.sheetIdMap(key)
	if err != nil {
		return err
	}
	sheetId, ok := sheetIdMap[srcName]
	if !ok {
		return fmt.Errorf("sheet_id not found. key:%s name:%s", key, srcName)
	}
	_, err = ss.service.Spreadsheets.BatchUpdate(key, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			&sheets.Request{
				DuplicateSheet: &sheets.DuplicateSheetRequest{
					NewSheetName:     dstName,
					SourceSheetId:    sheetId,
					InsertSheetIndex: int64(len(sheetIdMap)),
				},
			},
		},
	}).Do()
	if err != nil {
		return err
	}
	return nil
}

func (ss *Spreadsheet) SheetDelete(key, name string) error {
	sheetIdMap, err := ss.sheetIdMap(key)
	if err != nil {
		return err
	}
	sheetId, ok := sheetIdMap[name]
	if !ok {
		return fmt.Errorf("sheet_id not found. key:%s name:%s", key, name)
	}
	_, err = ss.service.Spreadsheets.BatchUpdate(key, &sheets.BatchUpdateSpreadsheetRequest{
		Requests: []*sheets.Request{
			&sheets.Request{
				DeleteSheet: &sheets.DeleteSheetRequest{
					SheetId: sheetId,
				},
			},
		},
	}).Do()
	if err != nil {
		return err
	}
	return nil
}

type Worksheet struct {
	service          *sheets.Service
	sheetKey         string
	sheetName        string
	values           [][]string
	headers          []string
	headerIndexes    []int
	Rows             []map[string]string
	MajorDimension   string
	ValueInputOption string
}

func (ws *Worksheet) SheetKey() string {
	return ws.sheetKey
}

func (ws *Worksheet) SheetName() string {
	return ws.sheetName
}

func (ws *Worksheet) Headers() []string {
	headers := []string{}
	for _, v := range ws.headers {
		if v != "" {
			headers = append(headers, v)
		}
	}
	return headers
}

func (ws *Worksheet) Refresh() error {
	r, err := ws.service.Spreadsheets.Values.Get(ws.sheetKey, ws.sheetName).Do()
	if err != nil {
		return err
	}
	if len(r.Values) <= 0 {
		return fmt.Errorf("no header. key:%s sheetName:%s", ws.sheetKey, ws.sheetName)
	}
	var (
		cols          = len(r.Values[0])
		headers       = make([]string, 0, cols)
		headerIndexes = make([]int, 0, cols)
		values        = make([][]string, 0, len(r.Values)-1)
	)
	for i, v := range r.Values[0] {
		if v != "" {
			headers = append(headers, v.(string))
			headerIndexes = append(headerIndexes, i)
		}
	}
	for _, vals := range r.Values[1:] {
		value := make([]string, cols)
		for i, v := range vals {
			if i == cols {
				break
			}
			value[i] = v.(string)
		}
		values = append(values, value)
	}
	ws.values = values
	ws.headers = headers
	ws.headerIndexes = headerIndexes
	ws.DiscardChanges()
	return nil
}

func (ws *Worksheet) Append(rows []map[string]string) error {
	var (
		v    = make([][]interface{}, len(rows))
		tmps = make([][]string, len(rows))
	)
	for i, row := range rows {
		t := make([]interface{}, len(ws.values[0]))
		u := make([]string, len(ws.values[0]))
		for j, hi := range ws.headerIndexes {
			h := ws.headers[j]
			if h != "" {
				t[hi] = row[h]
				u[hi] = row[h]
			} else {
				t[hi] = ""
				u[hi] = ""
			}
		}
		v[i] = t
		tmps[i] = u
	}
	_, err := ws.service.Spreadsheets.Values.Append(
		ws.sheetKey,
		fmt.Sprintf("%s!A%d", ws.sheetName, len(ws.values)+2),
		&sheets.ValueRange{
			MajorDimension: ws.MajorDimension,
			Values:         v,
		},
	).ValueInputOption(ws.ValueInputOption).Do()
	if err != nil {
		return err
	}
	ws.values = append(ws.values, tmps...)
	ws.DiscardChanges()
	return nil
}

func (ws *Worksheet) Values() [][]string {
	values := make([][]string, len(ws.values))
	for i, v := range ws.values {
		values[i] = make([]string, len(v))
		copy(values[i], v)
	}
	return values
}

func (ws *Worksheet) DiscardChanges() {
	rows := make([]map[string]string, 0, len(ws.values))
	for _, vals := range ws.values {
		row := make(map[string]string, len(vals))
		for i, headerIndex := range ws.headerIndexes {
			var (
				k = ws.headers[i]
				v = ""
			)
			if headerIndex < len(vals) {
				v = vals[headerIndex]
			}
			row[k] = v
		}
		rows = append(rows, row)
	}
	ws.Rows = rows
}

func (ws *Worksheet) Update() error {
	type tmp struct {
		row int
		col int
		val string
	}

	var (
		tmps = []tmp{}
		data = []*sheets.ValueRange{}
	)
	for r, row := range ws.Rows {
		for j, k := range ws.headers {
			c := ws.headerIndexes[j]
			v := row[k]
			if v != ws.values[r][c] {
				tmps = append(tmps, tmp{
					row: r,
					col: c,
					val: row[k],
				})
				data = append(data, &sheets.ValueRange{
					MajorDimension: ws.MajorDimension,
					Range: fmt.Sprintf(
						"%s!%s%d:%s%d",
						ws.sheetName, n2c(c+1), r+2, n2c(c+1), r+2,
					),
					Values: [][]interface{}{
						[]interface{}{v},
					},
				})
			}
		}
	}
	_, err := ws.service.Spreadsheets.Values.BatchUpdate(
		ws.sheetKey,
		&sheets.BatchUpdateValuesRequest{
			Data:             data,
			ValueInputOption: ws.ValueInputOption,
		},
	).Do()
	if err != nil {
		return err
	}
	for _, t := range tmps {
		ws.values[t.row][t.col] = t.val
	}
	return nil
}

func n2c(i int) string {
	j := 0
	r := ""
	for {
		i = i - 1
		j = i % 26
		i = i / 26
		if 0 < i {
			r = fmt.Sprintf("%s%s", string('A'+j), r)
		} else {
			break
		}
	}
	return fmt.Sprintf("%s%s", string('A'+j), r)
}
