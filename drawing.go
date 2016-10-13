package xlsx

/*

width 1 = 76200
width 80 = 640px

height 1 = 11430
height 75 = 100px

*/

type ImageType int

const (
	IMAGE_TYPE_JPG ImageType = iota
	IMAGE_TYPE_GIF
	IMAGE_TYPE_PNG
)

const (
	IMAGE_EXT_JPG = ".jpeg"
	IMAGE_EXT_GIF = ".gif"
	IMAGE_EXT_PNG = ".png"
)

const (
	PixelPerUnitWidth   = float64(8)
	PixelPerUnitHeight  = 100.0 / 75.0
	NumberPerUnitWidth  = float64(76200)
	NumberPerUnitHeight = float64(11430)
	UnitHeightPerCell   = float64(16.5)
)

type Drawing struct {
	Sheet       *Sheet
	ImageData   []byte
	ImageType   ImageType
	TopLeftCell DrawingCell
	RowCount    int
	ColCount    int
	Width       int
	Height      int
}

type DrawingCell struct {
	RowNum int
	ColNum int
}
