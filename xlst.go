package xlst

import (
	"errors"
	"io"
	"reflect"
	"regexp"

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

// New() creates new Xlst struct and returns pointer to it
func New() *Xlst {
	return &Xlst{}
}

// Render() renders report and stores it in a struct
func (m *Xlst) Render(ctx map[string]interface{}) error {
	report := xlsx.NewFile()
	for i, sheet := range m.file.Sheets {
		report.AddSheet(sheet.Name)
		cloneSheet(sheet, report.Sheets[i])
		for _, row := range sheet.Rows {
			prop := getListProp(row)
			if prop == "" {
				newRow := report.Sheets[0].AddRow()
				cloneRow(row, newRow)
				err := renderRow(newRow, ctx)
				if err != nil {
					return err
				}
				continue
			}
			if !isArray(ctx, prop) {
				newRow := report.Sheets[0].AddRow()
				cloneRow(row, newRow)
				err := renderRow(newRow, ctx)
				if err != nil {
					return err
				}
				continue
			}

			arr := reflect.ValueOf(ctx[prop])
			arrBackup := ctx[prop]
			for i := 0; i < arr.Len(); i++ {
				newRow := report.Sheets[0].AddRow()
				cloneRow(row, newRow)
				ctx[prop] = arr.Index(i).Interface()
				err := renderRow(newRow, ctx)
				if err != nil {
					return err
				}
			}
			ctx[prop] = arrBackup
		}
	}
	m.report = report

	return nil
}

// ReadTemplate() reads template from disk and stores it in a struct
func (m *Xlst) ReadTemplate(path string) error {
	file, err := xlsx.OpenFile(path)
	if err != nil {
		return err
	}
	m.file = file
	return nil
}

// Save() saves generated report to disk
func (m *Xlst) Save(path string) error {
	if m.report == nil {
		return errors.New("Report was not generated")
	}
	return m.report.Save(path)
}

// Write() writes generated report to provided writer
func (m *Xlst) Write(writer io.Writer) error {
	if m.report == nil {
		return errors.New("Report was not generated")
	}
	return m.report.Write(writer)
}

func cloneCell(from, to *xlsx.Cell) {
	to.Value = from.Value
	to.SetStyle(from.GetStyle())
	to.HMerge = from.HMerge
	to.VMerge = from.VMerge
	to.Hidden = from.Hidden
	to.NumFmt = from.NumFmt
}

func cloneRow(from, to *xlsx.Row) {
	to.Height = from.Height
	for _, cell := range from.Cells {
		newCell := to.AddCell()
		cloneCell(cell, newCell)
	}
}

func renderCell(cell *xlsx.Cell, ctx interface{}) error {
	template, err := raymond.Parse(cell.Value)
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
		newCol.SetStyle(col.GetStyle())
		newCol.Width = col.Width
		newCol.Hidden = col.Hidden
		newCol.Collapsed = col.Collapsed
		newCol.Min = col.Min
		newCol.Max = col.Max
		to.Cols = append(to.Cols, &newCol)
	}
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
