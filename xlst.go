package xlst

import (
	"errors"
	"fmt"
	"io"
	"reflect"
	"regexp"
	"strings"

	"github.com/aymerick/raymond"
	"github.com/tealeg/xlsx"
)

var (
	rgx         = regexp.MustCompile(`\{\{\s*(\w+)\.\w+\s*\}\}`)
	blockRgx    = regexp.MustCompile(`\{\{\s*(\w+)\s+(\w+)\s*\}\}`)
	blockEndRgx = regexp.MustCompile(`\{\{\s*end\s*\}\}`)
)

const (
	rangeBlock = "range"
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

		newSheet, err := report.AddSheet(sheet.Name)
		if err != nil {
			return fmt.Errorf("Cannot add sheet: %v", err)
		}

		cloneSheet(sheet, newSheet)

		sr := sheetRenderer{
			sheet:   newSheet,
			options: options,
		}

		if err := sr.render(sheet.Rows, ctx); err != nil {
			return err
		}

		newSheet.Cols = append(newSheet.Cols, sheet.Cols...)
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

type sheetRenderer struct {
	sheet   *xlsx.Sheet
	options *Options
}

func (r *sheetRenderer) render(rows []*xlsx.Row, ctx map[string]interface{}) error {
	var idx int

	for idx < len(rows) {
		b := getBlock(rows[idx])

		switch {

		case b != nil:
			blockLen, err := r.renderBlock(b, rows[idx:], ctx)
			if err != nil {
				return err
			}

			idx += blockLen

		default:
			err := r.renderRow(rows[idx], ctx)
			if err != nil {
				return err
			}

			idx++

		}
	}

	return nil
}

func (r *sheetRenderer) renderBlock(b *block, rows []*xlsx.Row, ctx map[string]interface{}) (int, error) {
	blockLen := getBlockLen(rows)
	if blockLen == -1 {
		return 0, fmt.Errorf("End of block {{%s %s}} not found", b.Name, b.Prop)
	}

	rangeCtx := getRangeCtx(ctx, b.Prop)
	if rangeCtx == nil {
		return 0, fmt.Errorf("Not expected context property for range %q", b.Prop)
	}

	for idx := range rangeCtx {
		localCtx := mergeCtx(rangeCtx[idx], ctx)
		err := r.render(rows[1:blockLen-1], localCtx)
		if err != nil {
			return 0, err
		}
	}

	return blockLen, nil
}

func (r *sheetRenderer) renderRow(row *xlsx.Row, ctx map[string]interface{}) error {
	prop := getListProp(row)
	if prop == "" || !isArray(ctx, prop) {
		newRow := r.sheet.AddRow()
		cloneRow(row, newRow, r.options)
		return renderRow(newRow, ctx)
	}

	arr := reflect.ValueOf(ctx[prop])
	arrBackup := ctx[prop]
	for i := 0; i < arr.Len(); i++ {
		newRow := r.sheet.AddRow()
		cloneRow(row, newRow, r.options)
		ctx[prop] = arr.Index(i).Interface()
		err := renderRow(newRow, ctx)
		if err != nil {
			return err
		}
	}
	ctx[prop] = arrBackup

	return nil
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
	if from.Height != 0 {
		to.SetHeight(from.Height)
	}

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

func getRangeCtx(ctx map[string]interface{}, prop string) []map[string]interface{} {
	val, ok := ctx[prop]
	if !ok {
		return nil
	}

	if propCtx, ok := val.([]map[string]interface{}); ok {
		return propCtx
	}

	return nil
}

func mergeCtx(local, global map[string]interface{}) map[string]interface{} {
	ctx := make(map[string]interface{})

	for k, v := range global {
		ctx[k] = v
	}

	for k, v := range local {
		ctx[k] = v
	}

	return ctx
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

		if match := rgx.FindStringSubmatch(cell.Value); match != nil {
			return match[1]
		}
	}
	return ""
}

func getBlockLen(rows []*xlsx.Row) int {
	var nesting int
	for idx := 1; idx < len(rows); idx++ {
		if len(rows[idx].Cells) == 0 {
			continue
		}

		if blockEndRgx.MatchString(rows[idx].Cells[0].Value) {
			if nesting == 0 {
				return idx + 1
			}

			nesting--
			continue
		}

		if blockRgx.MatchString(rows[idx].Cells[0].Value) {
			nesting++
		}
	}

	return -1
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

type block struct {
	Name string
	Prop string
}

func getBlock(row *xlsx.Row) *block {
	if len(row.Cells) == 0 || row.Cells[0].Value == "" {
		return nil
	}

	match := blockRgx.FindStringSubmatch(row.Cells[0].Value)
	if match == nil || match[1] != rangeBlock {
		return nil
	}

	return &block{
		Name: match[1],
		Prop: match[2],
	}
}
