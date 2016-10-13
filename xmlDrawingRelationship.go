package xlsx

import (
	"encoding/xml"
	"fmt"
)

type xlsxDrawingRelationships struct {
	XMLName       xml.Name                   `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []*xlsxDrawingRelationship ``
}

type xlsxDrawingRelationship struct {
	XMLName xml.Name `xml:"Relationship"`
	Id      string   `xml:",attr"`
	Type    string   `xml:",attr"`
	Target  string   `xml:",attr"`
}

func newXlsxDrawingRelationships() *xlsxDrawingRelationships {
	relationships := new(xlsxDrawingRelationships)
	relationships.Relationships = make([]*xlsxDrawingRelationship, 0)
	return relationships
}

func (relationships *xlsxDrawingRelationships) AddDrawingRelationship(imageName string) string {
	relationship := new(xlsxDrawingRelationship)
	relationship.Id = fmt.Sprintf("rId%d", len(relationships.Relationships)+1)
	relationship.Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	relationship.Target = fmt.Sprintf("../media/%s", imageName)
	relationships.Relationships = append(relationships.Relationships, relationship)
	return relationship.Id
}
