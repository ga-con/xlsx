package xlsx

import (
	"encoding/xml"
	"fmt"
)

type xlsxWorksheetRelationships struct {
	XMLName       xml.Name                     `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []*xlsxWorksheetRelationship ``
}

type xlsxWorksheetRelationship struct {
	XMLName xml.Name `xml:"Relationship"`
	Id      string   `xml:",attr"`
	Type    string   `xml:",attr"`
	Target  string   `xml:",attr"`
}

func newXlsxWorksheetRelationships() *xlsxWorksheetRelationships {
	relationships := new(xlsxWorksheetRelationships)
	relationships.Relationships = make([]*xlsxWorksheetRelationship, 0)
	return relationships
}

func (relationships *xlsxWorksheetRelationships) AddWorksheetDrawingRelationship(drawingXML string) string {
	relationship := new(xlsxWorksheetRelationship)
	relationship.Id = fmt.Sprintf("rId%d", len(relationships.Relationships)+1)
	relationship.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
	relationship.Target = fmt.Sprintf("../drawings/%s", drawingXML)
	relationships.Relationships = append(relationships.Relationships, relationship)
	return relationship.Id
}
