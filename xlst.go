package xlst

import (
	"errors"
	"fmt"
	"io"
	"reflect"
	"regexp"
	"strings"

	"github.com/aymerick/raymond"
	"github.com/kkbblzq/xlsx/v3"
)

var (
	rgx         = regexp.MustCompile(`\{\{\s*(\w+)\.\w+\s*\}\}`)
	rangeRgx    = regexp.MustCompile(`\{\{\s*range\s+(\w+)\s*\}\}`)
	rangeEndRgx = regexp.MustCompile(`\{\{\s*end\s*\}\}`)
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

func (m *Xlst) getRows(st *xlsx.Sheet) ([]*xlsx.Row, error) {
	rows := make([]*xlsx.Row, 0, st.MaxRow)
	if err := st.ForEachRow(func(r *xlsx.Row) error {
		rows = append(rows, r)
		return nil
	}); err != nil {
		return nil, err
	}

	return rows, nil
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
			return err
		}

		cloneSheet(sheet, newSheet)
		rows, err := m.getRows(sheet)
		if err != nil {
			return err
		}

		if err = renderRows(newSheet, rows, ctx, options); err != nil {
			return err
		}

		sheet.Cols.ForEach(func(idx int, col *xlsx.Col) {
			newSheet.Cols.Add(col)
		})
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

func renderRows(sheet *xlsx.Sheet, rows []*xlsx.Row, ctx map[string]interface{}, options *Options) error {
	for ri := 0; ri < len(rows); ri++ {
		row := rows[ri]

		rangeProp := getRangeProp(row)
		if rangeProp != "" {
			ri++

			rangeEndIndex := getRangeEndIndex(rows[ri:])
			if rangeEndIndex == -1 {
				return fmt.Errorf("End of range %q not found", rangeProp)
			}

			rangeEndIndex += ri

			rangeCtx := getRangeCtx(ctx, rangeProp)
			if rangeCtx == nil {
				return fmt.Errorf("Not expected context property for range %q", rangeProp)
			}

			for idx := range rangeCtx {
				localCtx := mergeCtx(rangeCtx[idx], ctx)
				err := renderRows(sheet, rows[ri:rangeEndIndex], localCtx, options)
				if err != nil {
					return err
				}
			}

			ri = rangeEndIndex

			continue
		}

		prop := getListProp(row)
		if prop == "" {
			newRow := sheet.AddRow()
			if err := cloneRow(row, newRow, options); err != nil {
				return err
			}

			if err := renderRow(newRow, ctx); err != nil {
				return err
			}
			continue
		}

		if !isArray(ctx, prop) {
			newRow := sheet.AddRow()
			if err := cloneRow(row, newRow, options); err != nil {
				return err
			}
			err := renderRow(newRow, ctx)
			if err != nil {
				return err
			}
			continue
		}

		arr := reflect.ValueOf(ctx[prop])
		arrBackup := ctx[prop]
		for i := 0; i < arr.Len(); i++ {
			newRow := sheet.AddRow()
			if err := cloneRow(row, newRow, options); err != nil {
				return err
			}
			ctx[prop] = arr.Index(i).Interface()
			err := renderRow(newRow, ctx)
			if err != nil {
				return err
			}
		}
		ctx[prop] = arrBackup
	}

	return nil
}

func cloneCell(from, to *xlsx.Cell, options *Options) error {
	fromBin, err := from.MarshalBinary()

	if err != nil {
		return err
	}

	if err := to.UnmarshalBinary(fromBin); err != nil {
		return err
	}

	style := from.GetStyle()
	if options.WrapTextInAllCells {
		style.Alignment.WrapText = true
	}

	to.SetStyle(style)
	if from.Value == "" && len(from.RichText) > 0 {
		to.SetRichText(from.RichText)
	}

	return nil
}

func cloneRow(from, to *xlsx.Row, options *Options) error {
	if from.GetHeight() != 0 {
		to.SetHeight(from.GetHeight())
	}

	return from.ForEachCell(func(cell *xlsx.Cell) error {
		newCell := to.AddCell()
		return cloneCell(cell, newCell, options)
	})
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
	from.Cols.ForEach(func(idx int, col *xlsx.Col) {
		newCol := xlsx.NewColForRange(col.Min, col.Max)
		style := col.GetStyle()
		newCol.SetStyle(style)
		newCol.Width = col.Width
		newCol.Hidden = col.Hidden
		newCol.Collapsed = col.Collapsed
		newCol.CustomWidth = col.CustomWidth
		newCol.OutlineLevel = col.OutlineLevel
		newCol.Phonetic = col.Phonetic
		to.Cols.Add(col)
	})
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
	for i := 0; i < in.Sheet.MaxCol; i++ {
		cell := in.GetCell(i)
		if cell.Value == "" {
			continue
		}
		if match := rgx.FindAllStringSubmatch(cell.Value, -1); match != nil {
			return match[0][1]
		}
	}
	return ""
}

func getRangeProp(in *xlsx.Row) string {
	if in.Sheet.MaxCol != 0 {
		match := rangeRgx.FindAllStringSubmatch(in.GetCell(0).Value, -1)
		if match != nil {
			return match[0][1]
		}
	}

	return ""
}

func getRangeEndIndex(rows []*xlsx.Row) int {
	var nesting int
	for idx := 0; idx < len(rows); idx++ {
		if rows[idx].Sheet.MaxCol == 0 {
			continue
		}

		if rangeEndRgx.MatchString(rows[idx].GetCell(0).Value) {
			if nesting == 0 {
				return idx
			}

			nesting--
			continue
		}

		if rangeRgx.MatchString(rows[idx].GetCell(0).Value) {
			nesting++
		}
	}

	return -1
}

func renderRow(in *xlsx.Row, ctx interface{}) error {
	return in.ForEachCell(func(cell *xlsx.Cell) error {
		err := renderCell(cell, ctx)
		if err != nil {
			return err
		}

		return nil
	})
}
