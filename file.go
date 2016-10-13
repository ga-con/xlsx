package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
)

// File is a high level structure providing a slice of Sheet structs
// to the user.
type File struct {
	worksheets     map[string]*zip.File
	referenceTable *RefTable
	Date1904       bool
	styles         *xlsxStyleSheet
	Sheets         []*Sheet
	Sheet          map[string]*Sheet
	theme          *theme
	DefinedNames   []*xlsxDefinedName
	Drawings       [][]Drawing
}

// Create a new File
func NewFile() *File {
	return &File{
		Sheet:        make(map[string]*Sheet),
		Sheets:       make([]*Sheet, 0),
		DefinedNames: make([]*xlsxDefinedName, 0),
		Drawings:     make([][]Drawing, 0),
	}
}

// OpenFile() take the name of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenFile(filename string) (file *File, err error) {
	var f *zip.ReadCloser
	f, err = zip.OpenReader(filename)
	if err != nil {
		return nil, err
	}
	file, err = ReadZip(f)
	return
}

// OpenBinary() take bytes of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenBinary(bs []byte) (*File, error) {
	r := bytes.NewReader(bs)
	return OpenReaderAt(r, int64(r.Len()))
}

// OpenReaderAt() take io.ReaderAt of an XLSX file and returns a populated
// xlsx.File struct for it.
func OpenReaderAt(r io.ReaderAt, size int64) (*File, error) {
	file, err := zip.NewReader(r, size)
	if err != nil {
		return nil, err
	}
	return ReadZipReader(file)
}

// A convenient wrapper around File.ToSlice, FileToSlice will
// return the raw data contained in an Excel XLSX file as three
// dimensional slice.  The first index represents the sheet number,
// the second the row number, and the third the cell number.
//
// For example:
//
//    var mySlice [][][]string
//    var value string
//    mySlice = xlsx.FileToSlice("myXLSX.xlsx")
//    value = mySlice[0][0][0]
//
// Here, value would be set to the raw value of the cell A1 in the
// first sheet in the XLSX file.
func FileToSlice(path string) ([][][]string, error) {
	f, err := OpenFile(path)
	if err != nil {
		return nil, err
	}
	return f.ToSlice()
}

// Save the File to an xlsx file at the provided path.
func (f *File) Save(path string) (err error) {
	target, err := os.Create(path)
	if err != nil {
		return err
	}
	err = f.Write(target)
	if err != nil {
		return err
	}
	return target.Close()
}

// Write the File to io.Writer as xlsx
func (f *File) Write(writer io.Writer) (err error) {
	parts, err := f.MarshallParts()
	if err != nil {
		return
	}
	zipWriter := zip.NewWriter(writer)
	for partName, part := range parts {
		w, err := zipWriter.Create(partName)
		if err != nil {
			return err
		}
		_, err = w.Write([]byte(part))
		if err != nil {
			return err
		}
	}
	return zipWriter.Close()
}

// Add a new Sheet, with the provided name, to a File
func (f *File) AddSheet(sheetName string) (*Sheet, error) {
	if _, exists := f.Sheet[sheetName]; exists {
		return nil, fmt.Errorf("duplicate sheet name '%s'.", sheetName)
	}
	sheet := &Sheet{
		Name:     sheetName,
		File:     f,
		Selected: len(f.Sheets) == 0,
	}
	f.Sheet[sheetName] = sheet
	f.Sheets = append(f.Sheets, sheet)
	return sheet, nil
}

func (f *File) makeWorkbook() xlsxWorkbook {
	return xlsxWorkbook{
		FileVersion: xlsxFileVersion{AppName: "Go XLSX"},
		WorkbookPr:  xlsxWorkbookPr{ShowObjects: "all"},
		BookViews: xlsxBookViews{
			WorkBookView: []xlsxWorkBookView{
				{
					ShowHorizontalScroll: true,
					ShowSheetTabs:        true,
					ShowVerticalScroll:   true,
					TabRatio:             204,
					WindowHeight:         8192,
					WindowWidth:          16384,
					XWindow:              "0",
					YWindow:              "0",
				},
			},
		},
		Sheets: xlsxSheets{Sheet: make([]xlsxSheet, len(f.Sheets))},
		CalcPr: xlsxCalcPr{
			IterateCount: 100,
			RefMode:      "A1",
			Iterate:      false,
			IterateDelta: 0.001,
		},
	}
}

// Some tools that read XLSX files have very strict requirements about
// the structure of the input XML.  In particular both Numbers on the Mac
// and SAS dislike inline XML namespace declarations, or namespace
// prefixes that don't match the ones that Excel itself uses.  This is a
// problem because the Go XML library doesn't multiple namespace
// declarations in a single element of a document.  This function is a
// horrible hack to fix that after the XML marshalling is completed.
func replaceRelationshipsNameSpace(workbookMarshal string) string {
	newWorkbook := strings.Replace(workbookMarshal, `xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships:id`, `r:id`, -1)
	// Dirty hack to fix issues #63 and #91; encoding/xml currently
	// "doesn't allow for additional namespaces to be defined in the
	// root element of the document," as described by @tealeg in the
	// comments for #63.
	oldXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
	newXmlns := `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
	return strings.Replace(newWorkbook, oldXmlns, newXmlns, 1)
}

// replaceWorksheetNameSpace print option issue
func replaceWorksheetNameSpace(worksheetMarshal string) string {
	oldXmlns := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
	newXmlns := `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
	return strings.Replace(worksheetMarshal, oldXmlns, newXmlns, 1)
}

// Construct a map of file name to XML content representing the file
// in terms of the structure of an XLSX file.
func (f *File) MarshallParts() (map[string]string, error) {
	var parts map[string]string
	var refTable *RefTable = NewSharedStringRefTable()
	refTable.isWrite = true
	var workbookRels WorkBookRels = make(WorkBookRels)
	var err error
	var workbook xlsxWorkbook
	var types xlsxTypes = MakeDefaultContentTypes()

	marshal := func(thing interface{}) (string, error) {
		body, err := xml.Marshal(thing)
		if err != nil {
			return "", err
		}

		outputStr := replaceWorksheetNameSpace(string(body))

		return strings.Replace(xml.Header, `<?xml version="1.0" encoding="UTF-8"?>`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`, -1) + outputStr, nil
	}

	parts = make(map[string]string)
	workbook = f.makeWorkbook()
	sheetIndex := 1
	drawingCount := 0

	if f.styles == nil {
		f.styles = newXlsxStyleSheet(f.theme)
	}
	f.styles.reset()

	for _, sheet := range f.Sheets {

		xSheet := sheet.makeXLSXSheet(refTable, f.styles)
		rId := fmt.Sprintf("rId%d", sheetIndex)
		sheetId := strconv.Itoa(sheetIndex)
		sheetPath := fmt.Sprintf("worksheets/sheet%d.xml", sheetIndex)
		partName := "xl/" + sheetPath
		types.Overrides = append(
			types.Overrides,
			xlsxOverride{
				PartName:    "/" + partName,
				ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"})
		workbookRels[rId] = sheetPath
		workbook.Sheets.Sheet[sheetIndex-1] = xlsxSheet{
			Name:    sheet.Name,
			SheetId: sheetId,
			Id:      rId,
			State:   "visible"}
		parts[partName], err = marshal(xSheet)

		if err != nil {
			return parts, err
		}

		xDrawing := newXlsxDrawing()
		xDrawingRel := newXlsxDrawingRelationships()

		colWidths := make([]float64, sheet.MaxCol)
		for _, col := range sheet.Cols {
			for i := col.Min - 1; i < col.Max; i++ {
				colWidths[i] = col.Width
			}
		}

		for _, drawing := range sheet.Drawings {
			drawingCount++
			var imageExt string
			switch drawing.ImageType {
			case IMAGE_TYPE_JPG:
				imageExt = IMAGE_EXT_JPG
			case IMAGE_TYPE_GIF:
				imageExt = IMAGE_EXT_GIF
			case IMAGE_TYPE_PNG:
				imageExt = IMAGE_EXT_PNG
			}
			imageName := fmt.Sprintf("image%d%s", drawingCount, imageExt)
			parts[fmt.Sprintf("xl/media/%s", imageName)] = string(drawing.ImageData)
			// TODO - calculate the bottom right cell location and offset
			var toCol, toColOff, toRow, toRowOff int
			if drawing.RowCount > 0 && drawing.ColCount > 0 {
				toCol = drawing.TopLeftCell.ColNum + drawing.ColCount
				toRow = drawing.TopLeftCell.RowNum + drawing.RowCount
			} else if drawing.RowCount > 0 {
				toRow = drawing.TopLeftCell.RowNum + drawing.RowCount
				targetHeightInPixel := PixelPerUnitHeight * UnitHeightPerCell * float64(drawing.RowCount)
				targetWidthInPixel := float64(drawing.Width) / float64(drawing.Height) * targetHeightInPixel
				targeWidth := targetWidthInPixel / PixelPerUnitWidth * NumberPerUnitWidth
				colIndex := drawing.TopLeftCell.ColNum
				for targeWidth >= colWidths[colIndex]*NumberPerUnitWidth {
					targeWidth -= colWidths[colIndex] * NumberPerUnitWidth
					colIndex++
				}
				toCol = colIndex
				toColOff = int(targeWidth)
			} else {
				toCol = drawing.TopLeftCell.ColNum + drawing.ColCount
				targetWidthInPixel := float64(0)
				for colIndex := drawing.TopLeftCell.ColNum; colIndex < toCol; colIndex++ {
					targetWidthInPixel += colWidths[colIndex] * PixelPerUnitWidth
				}
				targetHeightInPixel := float64(drawing.Height) / float64(drawing.Width) * targetWidthInPixel
				targetHeight := targetHeightInPixel / PixelPerUnitHeight * NumberPerUnitHeight
				rowIndex := drawing.TopLeftCell.RowNum
				fmt.Println(targetHeight)
				for targetHeight >= UnitHeightPerCell*NumberPerUnitHeight {
					targetHeight -= UnitHeightPerCell * NumberPerUnitHeight
					rowIndex++
				}
				toRow = rowIndex
				toRowOff = int(targetHeight)
				fmt.Println(targetHeight, rowIndex)
			}
			embedId := xDrawingRel.AddDrawingRelationship(imageName)
			xDrawing.AddDrawingTwoCellAnchor(drawing.TopLeftCell.ColNum, 0, drawing.TopLeftCell.RowNum, 0, toCol, toColOff, toRow, toRowOff, embedId)
		}

		drawingXML := fmt.Sprintf("drawing%d.xml", sheetIndex)
		drawingPartName := fmt.Sprintf("xl/drawings/%s", drawingXML)
		types.Overrides = append(
			types.Overrides,
			xlsxOverride{
				PartName:    "/" + drawingPartName,
				ContentType: "application/vnd.openxmlformats-officedocument.drawing+xml"})
		parts[fmt.Sprintf("xl/drawings/_rels/%s.rels", drawingXML)], err = marshal(xDrawingRel)
		if err != nil {
			return parts, err
		}
		parts[drawingPartName], err = marshal(xDrawing)
		if err != nil {
			return parts, err
		}
		xSheetRelationships := newXlsxWorksheetRelationships()
		xSheetRelationships.AddWorksheetDrawingRelationship(drawingXML)
		parts[fmt.Sprintf("xl/worksheets/_rels/sheet%d.xml.rels", sheetIndex)], err = marshal(xSheetRelationships)
		if err != nil {
			return parts, err
		}

		sheetIndex++
	}

	workbookMarshal, err := marshal(workbook)
	if err != nil {
		return parts, err
	}
	workbookMarshal = replaceRelationshipsNameSpace(workbookMarshal)
	parts["xl/workbook.xml"] = workbookMarshal
	if err != nil {
		return parts, err
	}

	parts["_rels/.rels"] = TEMPLATE__RELS_DOT_RELS
	parts["docProps/app.xml"] = TEMPLATE_DOCPROPS_APP
	// TODO - do this properly, modification and revision information
	parts["docProps/core.xml"] = TEMPLATE_DOCPROPS_CORE
	parts["xl/theme/theme1.xml"] = TEMPLATE_XL_THEME_THEME

	xSST := refTable.makeXLSXSST()
	parts["xl/sharedStrings.xml"], err = marshal(xSST)
	if err != nil {
		return parts, err
	}

	xWRel := workbookRels.MakeXLSXWorkbookRels()

	parts["xl/_rels/workbook.xml.rels"], err = marshal(xWRel)
	if err != nil {
		return parts, err
	}

	parts["[Content_Types].xml"], err = marshal(types)
	if err != nil {
		return parts, err
	}

	parts["xl/styles.xml"], err = f.styles.Marshal()
	if err != nil {
		return parts, err
	}

	return parts, nil
}

// Return the raw data contained in the File as three
// dimensional slice.  The first index represents the sheet number,
// the second the row number, and the third the cell number.
//
// For example:
//
//    var mySlice [][][]string
//    var value string
//    mySlice = xlsx.FileToSlice("myXLSX.xlsx")
//    value = mySlice[0][0][0]
//
// Here, value would be set to the raw value of the cell A1 in the
// first sheet in the XLSX file.
func (file *File) ToSlice() (output [][][]string, err error) {
	output = [][][]string{}
	for _, sheet := range file.Sheets {
		s := [][]string{}
		for _, row := range sheet.Rows {
			if row == nil {
				continue
			}
			r := []string{}
			for _, cell := range row.Cells {
				str, err := cell.String()
				if err != nil {
					return output, err
				}
				r = append(r, str)
			}
			s = append(s, r)
		}
		output = append(output, s)
	}
	return output, nil
}
