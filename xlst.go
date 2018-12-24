package xlst

import (
	"errors"
	"io"
	"reflect"
	"regexp"
	"strings"

	"github.com/aymerick/raymond"
	"github.com/tealeg/xlsx"
)

var (
	rgx = regexp.MustCompile(`\{\{\s*(\w+)\.\w+\s*\}\}`)
)

// Xlst Represents template struct
type Xlst struct {
	file   *xlsx.File
	report *xlsx.File
}

// Options for render has only one property WrapTextInAllCells for wrapping text
type Options struct {
	WrapTextInAllCells bool
}

// New creates new Xlst struct and returns pointer to it
func New() *Xlst {
	return &Xlst{}
}

// NewFromBinary creates new Xlst struct puts binary tempate into and returns pointer to it
func NewFromBinary(content []byte) (*Xlst, error) {
	file, err := xlsx.OpenBinary(content)
	if err != nil {
		return nil, err
	}
	
	res := &Xlst{file: file}
	return res, nil
}

// Render renders report and stores it in a struct
func (m *Xlst) Render(in interface{}) error {
	return m.RenderWithOptions(in, nil)
}

// RenderWithOptions renders report with options provided and stores it in a struct
func (m *Xlst) RenderWithOptions(in interface{}, options *Options) error {
	if options == nil {
		options = new(Options)
	}
	report := xlsx.NewFile()
	for si, sheet := range m.file.Sheets {
		ctx := getCtx(in, si)
		report.AddSheet(sheet.Name)
		cloneSheet(sheet, report.Sheets[si])
		for _, row := range sheet.Rows {
			prop := getListProp(row)
			if prop == "" {
				newRow := report.Sheets[si].AddRow()
				cloneRow(row, newRow, options)
				err := renderRow(newRow, ctx)
				if err != nil {
					return err
				}
				continue
			}
			if !isArray(ctx, prop) {
				newRow := report.Sheets[si].AddRow()
				cloneRow(row, newRow, options)
				err := renderRow(newRow, ctx)
				if err != nil {
					return err
				}
				continue
			}

			arr := reflect.ValueOf(ctx[prop])
			arrBackup := ctx[prop]
			for i := 0; i < arr.Len(); i++ {
				newRow := report.Sheets[si].AddRow()
				cloneRow(row, newRow, options)
				ctx[prop] = arr.Index(i).Interface()
				err := renderRow(newRow, ctx)
				if err != nil {
					return err
				}
			}
			ctx[prop] = arrBackup
		}
		for _, col := range sheet.Cols {
			report.Sheets[si].Cols = append(report.Sheets[si].Cols, col)
		}
	}
	m.report = report

	return nil
}

// ReadTemplate reads template from disk and stores it in a struct
func (m *Xlst) ReadTemplate(path string) error {
	file, err := xlsx.OpenFile(path)
	if err != nil {
		return err
	}
	m.file = file
	return nil
}

// Save saves generated report to disk
func (m *Xlst) Save(path string) error {
	if m.report == nil {
		return errors.New("Report was not generated")
	}
	return m.report.Save(path)
}

// Write writes generated report to provided writer
func (m *Xlst) Write(writer io.Writer) error {
	if m.report == nil {
		return errors.New("Report was not generated")
	}
	return m.report.Write(writer)
}

func cloneCell(from, to *xlsx.Cell, options *Options) {
	to.Value = from.Value
	style := from.GetStyle()
	if options.WrapTextInAllCells {
		style.Alignment.WrapText = true
	}
	to.SetStyle(style)
	to.HMerge = from.HMerge
	to.VMerge = from.VMerge
	to.Hidden = from.Hidden
	to.NumFmt = from.NumFmt
}

func cloneRow(from, to *xlsx.Row, options *Options) {
	to.Height = from.Height
	for _, cell := range from.Cells {
		newCell := to.AddCell()
		cloneCell(cell, newCell, options)
	}
}

func renderCell(cell *xlsx.Cell, ctx interface{}) error {
	tpl := strings.Replace(cell.Value, "{{", "{{{", -1)
	tpl = strings.Replace(tpl, "}}", "}}}", -1)
	template, err := raymond.Parse(tpl)
	if err != nil {
		return err
	}
	out, err := template.Exec(ctx)
	if err != nil {
		return err
	}
	cell.Value = out
	return nil
}

func cloneSheet(from, to *xlsx.Sheet) {
	for _, col := range from.Cols {
		newCol := xlsx.Col{}
		style := col.GetStyle()
		newCol.SetStyle(style)
		newCol.Width = col.Width
		newCol.Hidden = col.Hidden
		newCol.Collapsed = col.Collapsed
		newCol.Min = col.Min
		newCol.Max = col.Max
		to.Cols = append(to.Cols, &newCol)
	}
}

func getCtx(in interface{}, i int) map[string]interface{} {
	if ctx, ok := in.(map[string]interface{}); ok {
		return ctx
	}
	if ctxSlice, ok := in.([]interface{}); ok {
		if len(ctxSlice) > i {
			_ctx := ctxSlice[i]
			if ctx, ok := _ctx.(map[string]interface{}); ok {
				return ctx
			}
		}
		return nil
	}
	return nil
}

func isArray(in map[string]interface{}, prop string) bool {
	val, ok := in[prop]
	if !ok {
		return false
	}
	switch reflect.TypeOf(val).Kind() {
	case reflect.Array, reflect.Slice:
		return true
	}
	return false
}

func getListProp(in *xlsx.Row) string {
	for _, cell := range in.Cells {
		if cell.Value == "" {
			continue
		}
		if match := rgx.FindAllStringSubmatch(cell.Value, -1); match != nil {
			return match[0][1]
		}
	}
	return ""
}

func renderRow(in *xlsx.Row, ctx interface{}) error {
	for _, cell := range in.Cells {
		err := renderCell(cell, ctx)
		if err != nil {
			return err
		}
	}
	return nil
}
