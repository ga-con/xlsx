package xlsx

import (
	"encoding/xml"
)

type xlsxDrawing struct {
	XMLName        xml.Name                `xml:"xdr:wsDr"`
	NameSpace_XDR  string                  `xml:"xmlns:xdr,attr"`
	NameSpace_Main string                  `xml:"xmlns:a,attr"`
	TwoCellAnchors []*drawingTwoCellAnchor ``
}

type drawingTwoCellAnchor struct {
	XMLName    xml.Name          `xml:"xdr:twoCellAnchor"`
	EditAs     string            `xml:"editAs,attr"`
	From       drawingFrom       ``
	To         drawingTo         ``
	Pic        drawingPic        ``
	ClientData DrawingClientData ``
}

type drawingFrom struct {
	XMLName      xml.Name `xml:"xdr:from"`
	Column       int      `xml:"xdr:col"`
	ColumnOffset int      `xml:"xdr:colOff"`
	Row          int      `xml:"xdr:row"`
	RowOffset    int      `xml:"xdr:rowOff"`
}

type drawingTo struct {
	XMLName      xml.Name `xml:"xdr:to"`
	Column       int      `xml:"xdr:col"`
	ColumnOffset int      `xml:"xdr:colOff"`
	Row          int      `xml:"xdr:row"`
	RowOffset    int      `xml:"xdr:rowOff"`
}

type drawingPic struct {
	XMLName  xml.Name        `xml:"xdr:pic"`
	NvPicPr  drawingNvPicPr  ``
	BlipFill drawingBlipFill ``
	SpPr     drawingSpPr     ``
}

type drawingNvPicPr struct {
	XMLName  xml.Name        `xml:"xdr:nvPicPr"`
	CNvPr    drawingCNvPr    ``
	CNvPicPr drawingCNvPicPr ``
}

type drawingCNvPr struct {
	XMLName     xml.Name `xml:"xdr:cNvPr"`
	Id          int      `xml:"id,attr"`
	Name        string   `xml:"name,attr"`
	Description string   `xml:"descr,attr"`
}

type drawingCNvPicPr struct {
	XMLName  xml.Name     `xml:"xdr:cNvPicPr"`
	PicLocks mainPicLocks ``
}

type mainPicLocks struct {
	XMLName        xml.Name `xml:"a:picLocks"`
	NoChangeAspect int      `xml:"noChangeAspect,attr"`
}

type drawingBlipFill struct {
	XMLName xml.Name    `xml:"xdr:blipFill"`
	Blip    mainBlip    ``
	Stretch mainStretch ``
}

type mainBlip struct {
	XMLName     xml.Name `xml:"a:blip"`
	NameSpace_R string   `xml:"xmlns:r,attr"`
	Embed       string   `xml:"r:embed,attr"`
}

type mainStretch struct {
	XMLName  xml.Name     `xml:"a:stretch"`
	FillRect mainFillRect ``
}

type mainFillRect struct {
	XMLName xml.Name `xml:"a:fillRect"`
}

type drawingSpPr struct {
	XMLName  xml.Name     `xml:"xdr:spPr"`
	Xfrm     mainXfrm     ``
	PrstGeom mainPrstGeom ``
}

type mainXfrm struct {
	XMLName xml.Name `xml:"a:xfrm"`
	Off     mainOff  ``
	Ext     mainExt  ``
}

type mainOff struct {
	XMLName xml.Name `xml:"a:off"`
	X       int      `xml:"x,attr"`
	Y       int      `xml:"y,attr"`
}

type mainExt struct {
	XMLName xml.Name `xml:"a:ext"`
	CX      int      `xml:"cx,attr"`
	CY      int      `xml:"cy,attr"`
}

type mainPrstGeom struct {
	XMLName xml.Name  `xml:"a:prstGeom"`
	Prst    string    `xml:"prst,attr"`
	AvLst   mainAvLst ``
}

type mainAvLst struct {
	XMLName xml.Name `xml:"a:avLst"`
}

type DrawingClientData struct {
	XMLName xml.Name `xml:"xdr:clientData"`
}

func newXlsxDrawing() *xlsxDrawing {
	drawing := new(xlsxDrawing)
	drawing.NameSpace_Main = "http://schemas.openxmlformats.org/drawingml/2006/main"
	drawing.NameSpace_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
	drawing.TwoCellAnchors = make([]*drawingTwoCellAnchor, 0)
	return drawing
}

func (drawing *xlsxDrawing) AddDrawingTwoCellAnchor(fromCol, fromColOff, fromRow, fromRowOff, toCol, toColOff, toRow, toRowOff int, embedId string) {
	anchor := new(drawingTwoCellAnchor)
	anchor.EditAs = "oneCell"
	anchor.From.Column = fromCol
	anchor.From.ColumnOffset = fromColOff
	anchor.From.Row = fromRow
	anchor.From.RowOffset = fromRowOff
	anchor.To.Column = toCol
	anchor.To.ColumnOffset = toColOff
	anchor.To.Row = toRow
	anchor.To.RowOffset = toRowOff
	anchor.Pic.NvPicPr.CNvPr.Id = 0
	anchor.Pic.NvPicPr.CNvPicPr.PicLocks.NoChangeAspect = 1
	anchor.Pic.BlipFill.Blip.NameSpace_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	anchor.Pic.BlipFill.Blip.Embed = embedId
	anchor.Pic.SpPr.PrstGeom.Prst = "rect"
	drawing.TwoCellAnchors = append(drawing.TwoCellAnchors, anchor)
}
