package gss

import (
	"bytes"
	"encoding/json"
	"io/ioutil"
	"net/http"
)

type mockTransport struct {
	index             int
	dummyJsonResponse []interface{}
	req               []*http.Request
}

func (t *mockTransport) RoundTrip(req *http.Request) (*http.Response, error) {
	t.req = append(t.req, req)
	res := &http.Response{
		Header:     make(http.Header),
		Request:    req,
		StatusCode: http.StatusOK,
	}
	res.Header.Set("Content-Type", "application/json")
	b, err := json.Marshal(t.dummyJsonResponse[t.index])
	if err != nil {
		return nil, err
	}
	res.Body = ioutil.NopCloser(bytes.NewReader(b))
	t.index = t.index + 1
	return res, nil
}

func newDummyClient(d ...interface{}) (*http.Client, *mockTransport) {
	m := &mockTransport{
		dummyJsonResponse: d,
		req:               make([]*http.Request, 0, len(d)),
	}
	return &http.Client{Transport: m}, m
}

func newDummyWorksheet() (*Worksheet, error) {
	m := &mockTransport{
		dummyJsonResponse: []interface{}{
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
		},
	}
	client := &http.Client{Transport: m}
	ss, err := NewSpreadsheet(client)
	if err != nil {
		return nil, err
	}
	return ss.GetWorksheet("XXXXXX", "シート1")
}
