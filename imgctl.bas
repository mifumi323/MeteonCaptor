Attribute VB_Name = "imgctl"
'---------------------------------------------------------------
'	[imgctl.bas] version 1.24
'		For imgctl.dll version 1.24.
'		Copyright (C) ruche.
'---------------------------------------------------------------

'---------------------------------------------------------------
' Module file support version

Public Const IMGCTL_VERSION = 124%
Public Const IMGCTL_BETA = 0%
Public Const IMGCTL_VERSION_STRING = "1.24"

'---------------------------------------------------------------
' PNGOPT flag Constants

Public Const POF_COMPLEVEL = &H00000001&
Public Const POF_FILTER = &H00000002&
Public Const POF_GAMMA = &H00000004&
Public Const POF_TRNSCOLOR = &H00000008&
Public Const POF_BACKCOLOR = &H00000010&
Public Const POF_TEXT = &H00000020&
Public Const POF_TEXTCOMP = &H00000040&
Public Const POF_INTERLACING = &H00000080&
Public Const POF_TIME = &H00000100&
Public Const POF_BACKPALETTE = &H00010000&
Public Const POF_TRNSPALETTE = &H00020000&
Public Const POF_TRNSALPHA = &H00040000&

'---------------------------------------------------------------
' PNGOPT filter Constants

Public Const PO_FILTER_NONE = &H08&
Public Const PO_FILTER_SUB = &H10&
Public Const PO_FILTER_UP = &H20&
Public Const PO_FILTER_AVG = &H40&
Public Const PO_FILTER_PAETH = &H80&
Public Const PO_FILTER_ALL = PO_FILTER_NONE Or PO_FILTER_SUB _
	Or PO_FILTER_UP Or PO_FILTER_AVG Or PO_FILTER_PAETH

'---------------------------------------------------------------
' PNGOPT gamma Constants

Public Const PO_GAMMA_NORMAL = 45455&
Public Const PO_GAMMA_WIN = PO_GAMMA_NORMAL
Public Const PO_GAMMA_MAC = 55556&

'---------------------------------------------------------------
' GIFOPT flag Constants (v1.13B4)

Public Const GOF_LOGICAL = &H00000001&
Public Const GOF_TRNSCOLOR = &H00000008&
Public Const GOF_BACKCOLOR = &H00000010&
Public Const GOF_INTERLACING = &H00000080&
Public Const GOF_BACKPALETTE = &H00010000&
Public Const GOF_TRNSPALETTE = &H00020000&
Public Const GOF_LZWCLRCOUNT = &H00080000&
Public Const GOF_LZWNOTUSE = &H00100000&
Public Const GOF_BITCOUNT = &H00200000&

'---------------------------------------------------------------
' GIFANIOPT flag Constants (v1.13)

Public Const GAF_LOGICAL = &H00000001&
Public Const GAF_BACKCOLOR = &H00000010&
Public Const GAF_LOOPCOUNT = &H00000400&
Public Const GAF_NOTANI = &H00000800&

'---------------------------------------------------------------
' GIFANISCENE flag Constants (v1.13)

Public Const GSF_LOGICAL = &H00000001&
Public Const GSF_TRNSCOLOR = &H00000008&
Public Const GSF_INTERLACING = &H00000080&
Public Const GSF_TRNSPALETTE = &H00020000&
Public Const GSF_LZWCLRCOUNT = &H00080000&
Public Const GSF_LZWNOTUSE = &H00100000&
Public Const GSF_BITCOUNT = &H00200000&
Public Const GSF_DISPOSAL = &H00001000&
Public Const GSF_USERINPUT = &H00002000&

'---------------------------------------------------------------
' GIFANISCENE disposal methods Constants (v1.13)

Public Const GS_DISP_NONE = 0%
Public Const GS_DISP_LEAVE = 1%
Public Const GS_DISP_BACK = 2%
Public Const GS_DISP_PREV = 3%

'---------------------------------------------------------------
' Image type Constants

Public Const IMG_ERROR = &H00FFFFFF&	' v1.12B9
Public Const IMG_UNKNOWN = &H00000000&
Public Const IMG_BMP = &H00000001&
Public Const IMG_BITMAP = IMG_BMP
Public Const IMG_DIB = IMG_BMP
Public Const IMG_RLE = IMG_BMP
Public Const IMG_JPEG = &H00000002&
Public Const IMG_JPE = IMG_JPEG
Public Const IMG_JPG = IMG_JPEG
Public Const IMG_EXIF = IMG_JPEG		' v1.10B4
Public Const IMG_PNG = &H00000003&
Public Const IMG_GIF = &H00000004&
Public Const IMG_TIFF = &H00000005&
Public Const IMG_TIF = IMG_TIFF
Public Const IMG_PIC = &H00000006&
Public Const IMG_MAG = &H00000007&
Public Const IMG_MAKI = IMG_MAG
Public Const IMG_PCX = &H00000008&

'---------------------------------------------------------------
' Enough buffer size (v1.16B3)

Public Const BUFSIZE_ENOUGH = &HFFFFFFFF&

'---------------------------------------------------------------
' Bitfield types

Public Const RGB555_R = &H00007C00&
Public Const RGB555_G = &H000003E0&
Public Const RGB555_B = &H0000001F&
Public Const RGB565_R = &H0000F800&
Public Const RGB565_G = &H000007E0&
Public Const RGB565_B = &H0000001F&
Public Const RGB888_R = &H00FF0000&
Public Const RGB888_G = &H0000FF00&
Public Const RGB888_B = &H000000FF&

'---------------------------------------------------------------
' DIBto16BitEx & DIBto8Bit types (v1.10B4)

Public Const TOBIT_DEFAULT = &H00000000&
Public Const TOBIT_ORG = &H00100000&		' v1.10
Public Const TOBIT_DIFF = &H00000001&
Public Const TOBIT_DIFFFS = &H00000002&		' v1.10B5 default
Public Const TOBIT_DIFFJJN = &H00000003&	' v1.10B5
Public Const TOBIT_DIFFX = &H00000101&		' v1.10B7
Public Const TOBIT_DIFFXFS = &H00000102&	' v1.10B7
Public Const TOBIT_DIFFXJJN = &H00000103&	' v1.10B7
Public Const TOBIT_DIFFDX = &H00000201&		' v1.10B8
Public Const TOBIT_DIFFDXFS = &H00000202&	' v1.10B8
Public Const TOBIT_DIFFDXJJN = &H00000203&	' v1.10B8

'---------------------------------------------------------------
' DIBto8Bit flags

Public Const TO8_DIV_RGB = &H00&
Public Const TO8_DIV_LIGHT = &H01&
Public Const TO8_SEL_CENTER = &H0000&
Public Const TO8_SEL_AVGRGB = &H0100&
Public Const TO8_SEL_AVGPIX = &H0200&
Public Const TO8_PUT_RGB = &H000000&
Public Const TO8_PUT_LIGHT = &H010000&
Public Const TO8_PUT_YUV = &H020000&

'---------------------------------------------------------------
' Resize flags (v1.12B)

Public Const RESZ_SAME = 0&
Public Const RESZ_RATIO = -1&

'---------------------------------------------------------------
' Replace colors

Public Const REP_R = 0%
Public Const REP_RED = REP_R
Public Const REP_G = 1%
Public Const REP_GREEN = REP_G
Public Const REP_B = 2%
Public Const REP_BLUE = REP_B

'---------------------------------------------------------------
' Turn types

Public Const TURN_90 = 90&
Public Const TURN_180 = 180&
Public Const TURN_270 = 270&

'---------------------------------------------------------------
' Error codes (v1.07B2)

Public Const ICERR_NONE = &H00000000&
Public Const ICERR_FILE_OPEN = &H00010001&
Public Const ICERR_FILE_READ = &H00010002&
Public Const ICERR_FILE_WRITE = &H00010003&
Public Const ICERR_FILE_TYPE = &H00010004&		' v1.07B3
Public Const ICERR_FILE_NULL = &H00010005&		' v1.07B3
Public Const ICERR_FILE_SEEK = &H00010006&		' v1.13B4
Public Const ICERR_FILE_SIZE = &H00010007&		' v1.16B3
Public Const ICERR_PARAM_NULL = &H00020001&
Public Const ICERR_PARAM_SIZE = &H00020002&
Public Const ICERR_PARAM_TYPE = &H00020003&
Public Const ICERR_PARAM_RANGE = &H00020004&	' v1.13
Public Const ICERR_MEM_ALLOC = &H00030001&
Public Const ICERR_MEM_SIZE = &H00030002&		' v1.16B3
Public Const ICERR_IMG_COMPRESS = &H00040001&
Public Const ICERR_IMG_RLESIZE = &H00040002&
Public Const ICERR_IMG_BITCOUNT = &H00040003&	' v1.12B3
Public Const ICERR_IMG_AREA = &H00040004&		' v1.12B3
Public Const ICERR_IMG_RLETOP = &H00040005&		' v1.12B3
Public Const ICERR_DIB_RLECOMP = &H00050001&
Public Const ICERR_DIB_RLEEXP = &H00050002&
Public Const ICERR_DIB_RLEBIT = &H00050003&
Public Const ICERR_DIB_NULL = &H00050004&		' v1.07B3
Public Const ICERR_DIB_UPPER16 = &H00050005&	' v1.07B3
Public Const ICERR_DIB_AREAOUT = &H00050006&	' v1.10B
Public Const ICERR_BMP_FILEHEAD = &H00060001&
Public Const ICERR_BMP_HEADSIZE = &H00060002&
Public Const ICERR_BMP_IMGSIZE = &H00060003&
Public Const ICERR_BMP_COMPRESS = &H00060004&
Public Const ICERR_RLE_TOPDOWN = &H00070001&
Public Const ICERR_RLE_DATASIZE = &H00070002&
Public Const ICERR_JPEG_LIBERR = &H00080001&
Public Const ICERR_PNG_LIBERR = &H00090001&
Public Const ICERR_PNG_NOALPHA = &H00090002&	' v1.12B5
Public Const ICERR_GIF_FILEHEAD = &H000B0001&	' v1.13B4
Public Const ICERR_GIF_BLOCK = &H000B0002&		' v1.13B4
Public Const ICERR_API_STRETCH = &H000A0001&
Public Const ICERR_API_SETMODE = &H000A0002&
Public Const ICERR_API_SECTION = &H000A0003&	' v1.09
Public Const ICERR_API_COMDC = &H000A0004&		' v1.09
Public Const ICERR_API_SELOBJ = &H000A0005&		' v1.09
Public Const ICERR_API_BITBLT = &H000A0006&		' v1.09

'---------------------------------------------------------------
' DIB compression Constants

Public Const BI_RGB = 0&
Public Const BI_RLE8 = 1&
Public Const BI_RLE4 = 2&
Public Const BI_BITFIELDS = 3&

'---------------------------------------------------------------
' Raster operation Constants

Public Const BLACKNESS = &H42&			' BLACK
Public Const DSTINVERT = &H550009&		' !dest
Public Const MERGECOPY = &HC000CA&		' src & pat
Public Const MERGEPAINT = &HBB0226&		' !src | dest
Public Const NOTSRCCOPY = &H330008&		' !src
Public Const NOTSRCERASE = &H1100A6&	' !src & !dest
Public Const PATCOPY = &HF00021&		' pat
Public Const PATINVERT = &H5A0049&		' pat ^ dest
Public Const PATPAINT = &HFB0A09&		' (!src | pat) | dest
Public Const SRCAND = &H8800C6&			' src & dest
Public Const SRCCOPY = &HCC0020&		' src
Public Const SRCERASE = &H440328&		' src & !dest
Public Const SRCINVERT = &H660046&		' src ^ dest
Public Const SRCPAINT = &HEE0086&		' src | dest
Public Const WHITENESS = &HFF0062&		' WHITE

'---------------------------------------------------------------
' Stretch mode Constants (v1.09)

Public Const BLACKONWHITE = 1&
Public Const WHITEONBLACK = 2&
Public Const COLORONCOLOR = 3&
Public Const HALFTONE = 4&

'---------------------------------------------------------------
' Structures

Public Type BITMAPINFOHEADER	' 40 bytes
	biSize As Long
	biWidth As Long
	biHeight As Long
	biPlanes As Integer
	biBitCount As Integer
	biCompression As Long
	biSizeImage As Long
	biXPelsPerMeter As Long
	biYPelsPerMeter As Long
	biClrUsed As Long
	biClrImportant As Long
End Type

Public Type RGBQUAD				' 4 bytes
	rgbBlue As Byte
	rgbGreen As Byte
	rgbRed As Byte
	rgbReserved As Byte
End Type

Public Type BITMAPINFO			' 44 ~ 1064 bytes
	bmiHeader As BITMAPINFOHEADER
	bmiColors(0 To 255) As RGBQUAD
End Type

Public Type RECT				' 16 bytes
	left As Long
	top As Long
	right As Long
	bottom As Long
End Type

'Public Type IMGDATA			' 16 bytes
'	dwbmiSize As Long
'	dwDataSize As Long
'	pbmi As Long
'	pData As Long
'End Type

Public Type PASTEINFO			' 16 bytes
	colDest As Long
	colSrc As Long
	dwReserved As Long
	lParam As Long

	' v1.10B
	rcArea As RECT
	lXDest As Long
	lYDest As Long
	lXSrc As Long
	lYSrc As Long
End Type

Public Type REPAINTINFO			' 8 bytes
	colBefore As Long
	colAfter As Long
End Type

Public Type CONVTABLE			' 768 bytes
	tblB(0 To 255) As Byte
	tblG(0 To 255) As Byte
	tblR(0 To 255) As Byte
End Type

Public Type PNGOPT				' 36 bytes
	dwFlag As Long
	wCompLevel As Integer

	wReserved As Integer	' v1.12B7

	dwFilter As Long
	dwGamma As Long
	clrTrans As Long
	clrBack As Long
	lpText As Long
	
	dwReserved1 As Long
	dwReserved2 As Long
End Type

Public Type PALTRANS			' 260 bytes
	trans(0 To 255) As Byte
	dwNum As Long
End Type

Public Type GIFOPT				' 28 bytes
	dwFlag As Long

	clrTrans As Long
	clrBack As Long

	wLogWidth As Integer
	wLogHeight As Integer
	wLogLeft As Integer
	wLogTop As Integer

	dwLzwCount As Long
	dwBitCount As Long
End Type

Public Type GIFANIOPT			' 16 bytes
	dwFlag As Long

	clrBack As Long
	wLogWidth As Integer
	wLogHeight As Integer
	wLoopCount As Integer

	wReserved As Integer
End Type

Public Type GIFANISCENE			' 28 bytes
	dwFlag As Long

	clrTrans As Long
	wLogLeft As Integer
	wLogTop As Integer
	dwLzwCount As Long
	dwBitCount As Long
	wDisposal As Integer

	wTime As Integer
	hDIB As Long
End Type

'---------------------------------------------------------------
' Standard Functions

Public Declare Function ImgctlVersion Lib "imgctl" ( _
	) As Integer
Public Declare Function ImgctlBeta Lib "imgctl" ( _
	) As Integer
Public Declare Function ImgctlError Lib "imgctl" ( _
	) As Long
Public Declare Sub ImgctlErrorClear Lib "imgctl" ()
Public Declare Function PointerOf Lib "imgctl" ( _
	ByRef pvData As Any _
	) As Long
Public Declare Function GetImageType Lib "imgctl" ( _
	ByVal lpImageFile As String, _
	ByRef pdwFlag As Long _
	) As Long
Public Declare Function GetImageMType Lib "imgctl" ( _
	ByRef pBuffer As Byte, _
	ByVal dwBufSize As Long, _
	ByRef pdwFlag As Long _
	) As Long
Public Declare Function ToDIB Lib "imgctl" ( _
	ByVal lpImageFile As String _
	) As Long	' HDIB
Public Declare Function MtoDIB Lib "imgctl" ( _
	ByRef pBuffer As Byte, _
	ByVal dwBufSize As Long _
	) As Long	' HDIB

'---------------------------------------------------------------
' DIB Functions

Public Declare Function DeleteDIB Lib "imgctl" ( _
	ByVal hDIB As Long _
	) As Long
Public Declare Function HeadDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByRef pbmih As BITMAPINFOHEADER _
	) As Long
Public Declare Function PaletteDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByRef rgbColors As RGBQUAD, _
	ByVal dwClrNum As Long _
	) As Long
Public Declare Function PixelDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal lXPos As Long, _
	ByVal lYPos As Long _
	) As Long
Public Declare Function ColorDIB Lib "imgctl" ( _
	ByVal hDIB As Long _
	) As Long
Public Declare Function GetDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByRef pbmi As BITMAPINFO, _
	ByRef pdwbmiSize As Long, _
	ByRef pvData As Any, _
	ByRef pdwDataSize As Long _
	) As Long
Public Declare Function MapDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByRef ppbmi As Long, _
	ByRef pdwbmiSize As Long, _
	ByRef ppvData As Long, _
	ByRef pdwDataSize As Long _
	) As Long
Public Declare Function DataDIB Lib "imgctl" ( _
	ByVal hDIB As Long _
	) As Long
Public Declare Function CreateDIB Lib "imgctl" ( _
	ByRef pbmi As BITMAPINFO, _
	ByRef pvData As Byte _
	) As Long	' HDIB
Public Declare Function CopyDIB Lib "imgctl" ( _
	ByVal hDIB As Long _
	) As Long	' HDIB
Public Declare Function CutDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal lX As Long, _
	ByVal lY As Long, _
	ByVal lWidth As Long, _
	ByVal lHeight As Long _
	) As Long	' HDIB
Public Declare Function TurnDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal iTurnType As Long _
	) As Long
Public Declare Function DIBto24Bit Lib "imgctl" ( _
	ByVal hDIB As Long _
	) As Long
Public Declare Function DIBto16Bit Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByRef dwBitFields As Long _
	) As Long
Public Declare Function DIBto16BitEx Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByRef dwBitFields As Long, _
	ByRef dwType As Long _
	) As Long
Public Declare Function DIBto8Bit Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal dwFlags As Long, _
	ByVal dwType As Long _
	) As Long

' 24Bit DIB Functions
Public Declare Function PasteDIB Lib "imgctl" ( _
	ByVal hDIBDest As Long, _
	ByVal lXDest As Long, _
	ByVal lYDest As Long, _
	ByVal lWidth As Long, _
	ByVal lHeight As Long, _
	ByVal hDIBSrc As Long, _
	ByVal lXSrc As Long, _
	ByVal lYSrc As Long, _
	ByVal pfnPasteProc As Long, _
	ByVal lParam As Long _
	) As Long
Public Declare Function ResizeDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal lWidth As Long, _
	ByVal lHeight As Long _
	) As Long
Public Declare Function TurnDIBex Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal lAngle As Long, _
	ByVal clrBack As Long _
	) As Long

'---------------------------------------------------------------
' RLE-DIB Functions

Public Declare Function IsRLE Lib "imgctl" ( _
	ByVal hDIB As Long _
	) As Long
Public Declare Function DIBtoRLE Lib "imgctl" ( _
	ByVal hDIB As Long _
	) As Long
Public Declare Function RLEtoDIB Lib "imgctl" ( _
	ByVal hRLE As Long _
	) As Long

'---------------------------------------------------------------
' Bitmap Functions

Public Declare Function DIBtoBMP Lib "imgctl" ( _
	ByVal lpBmpFile As String, _
	ByVal hDIB As Long _
	) As Long
Public Declare Function BMPtoDIB Lib "imgctl" ( _
	ByVal lpBmpFile As String _
	) As Long	' HDIB
Public Declare Function BMPMtoDIB Lib "imgctl" ( _
	ByRef pBuffer As Byte, _
	ByVal dwBufSize As Long _
	) As Long	' HDIB

'---------------------------------------------------------------
' JPEG Functions

Public Declare Function DIBtoJPG Lib "imgctl" ( _
	ByVal lpJpegFile As String, _
	ByVal hDIB As Long, _
	ByVal iQuality As Long, _
	ByVal bProgression As Long _
	) As Long
Public Declare Function JPGtoDIB Lib "imgctl" ( _
	ByVal lpJpegFile As String _
	) As Long	' HDIB
Public Declare Function JPGMtoDIB Lib "imgctl" ( _
	ByRef pBuffer As Byte, _
	ByVal dwBufSize As Long _
	) As Long	' HDIB

'---------------------------------------------------------------
' PNG Functions

Public Declare Function DIBtoPNG Lib "imgctl" ( _
	ByVal lpPngFile As String, _
	ByVal hDIB As Long, _
	ByVal bInterlacing As Long _
	) As Long
Public Declare Function DIBtoPNGex Lib "imgctl" ( _
	ByVal lpPngFile As String, _
	ByVal hDIB As Long, _
	ByRef pPngOpt As PNGOPT _
	) As Long
Public Declare Function PNGtoDIB Lib "imgctl" ( _
	ByVal lpPngFile As String _
	) As Long	' HDIB
Public Declare Function PNGMtoDIB Lib "imgctl" ( _
	ByRef pBuffer As Byte, _
	ByVal dwBufSize As Long _
	) As Long	' HDIB
Public Declare Function PNGAtoDIB Lib "imgctl" ( _
	ByVal lpPngFile As String _
	) As Long	' HDIB
Public Declare Function PNGMAtoDIB Lib "imgctl" ( _
	ByRef pBuffer As Byte, _
	ByVal dwBufSize As Long _
	) As Long	' HDIB
Public Declare Function InfoPNG Lib "imgctl" ( _
	ByVal lpPngFile As String, _
	ByRef pPngOpt As PNGOPT, _
	ByRef pbmi As BITMAPINFO, _
	ByVal dwbmiSize As Long _
	) As Long
Public Declare Function InfoPNGM Lib "imgctl" ( _
	ByRef pBuffer As Byte, _
	ByVal dwBufSize As Long, _
	ByRef pPngOpt As PNGOPT, _
	ByRef pbmi As BITMAPINFO, _
	ByVal dwbmiSize As Long _
	) As Long	' HDIB

'---------------------------------------------------------------
' GIF Functions

Public Declare Function DIBtoGIF Lib "imgctl" ( _
	ByVal lpGifFile As String, _
	ByVal hDIB As Long, _
	ByVal bInterlacing As Long _
	) As Long
Public Declare Function DIBtoGIFex Lib "imgctl" ( _
	ByVal lpGifFile As String, _
	ByVal hDIB As Long, _
	ByRef pGifOpt As GIFOPT _
	) As Long
Public Declare Function DIBtoGIFAni Lib "imgctl" ( _
	ByVal lpGifFile As String, _
	ByRef hDIBs As Long, _
	ByVal dwCount As Long, _
	ByVal wTime As Integer _
	) As Long
Public Declare Function DIBtoGIFAniEx Lib "imgctl" ( _
	ByVal lpGifFile As String, _
	ByRef aniScenes As GIFANISCENE, _
	ByVal dwCount As Long, _
	ByRef pGifAniOpt As GIFANIOPT _
	) As Long
Public Declare Function GIFtoDIB Lib "imgctl" ( _
	ByVal lpGifFile As String _
	) As Long	' HDIB
Public Declare Function GIFMtoDIB Lib "imgctl" ( _
	ByRef pBuffer As Byte, _
	ByVal dwBufSize As Long _
	) As Long	' HDIB
Public Declare Function GIFtoDIBex Lib "imgctl" ( _
	ByVal lpGifFile As String, _
	ByRef pGifOpt As GIFOPT _
	) As Long	' HDIB
Public Declare Function GIFMtoDIBex Lib "imgctl" ( _
	ByRef pBuffer As Byte, _
	ByVal dwBufSize As Long, _
	ByRef pGifOpt As GIFOPT _
	) As Long	' HDIB

'---------------------------------------------------------------
' Filter Functions

Public Declare Function GrayDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal wStrength As Integer _
	) As Long
Public Declare Function ReplaceDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal rTo As Integer, _
	ByVal gTo As Integer, _
	ByVal bTo As Integer _
	) As Long
Public Declare Function RepaintDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByRef repis As REPAINTINFO, _
	ByVal dwRepaintNum As Long _
	) As Long
Public Declare Function TableDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByRef ptable As CONVTABLE _
	) As Long

' Convert table functions
Public Declare Function ToneDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal rMove As Integer, _
	ByVal gMove As Integer, _
	ByVal bMove As Integer _
	) As Long
Public Declare Function ShadeDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal rMul As Long, _
	ByVal gMul As Long, _
	ByVal bMul As Long _
	) As Long
Public Declare Function GammaDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal rGma As Integer, _
	ByVal gGma As Integer, _
	ByVal bGma As Integer _
	) As Long
Public Declare Function ContrastDIB Lib "imgctl" ( _
	ByVal hDIB As Long, _
	ByVal rStr As Integer, _
	ByVal gStr As Integer, _
	ByVal bStr As Integer _
	) As Long

'---------------------------------------------------------------
' Device context Functions

Public Declare Function DIBtoDC Lib "imgctl" ( _
	ByVal hdc As Long, _
	ByVal nXDest As Long, _
	ByVal nYDest As Long, _
	ByVal nWidth As Long, _
	ByVal nHeight As Long, _
	ByVal hDIB As Long, _
	ByVal nXSrc As Long, _
	ByVal nYSrc As Long, _
	ByVal dwRop As Long _
	) As Long
Public Declare Function DIBtoDCex Lib "imgctl" ( _
	ByVal hdc As Long, _
	ByVal nXDest As Long, _
	ByVal nYDest As Long, _
	ByVal nDestWidth As Long, _
	ByVal nDestHeight As Long, _
	ByVal hDIB As Long, _
	ByVal nXSrc As Long, _
	ByVal nYSrc As Long, _
	ByVal nSrcWidth As Long, _
	ByVal nSrcHeight As Long, _
	ByVal dwRop As Long _
	) As Long
Public Declare Function DIBtoDCex2 Lib "imgctl" ( _
	ByVal hdc As Long, _
	ByVal nXDest As Long, _
	ByVal nYDest As Long, _
	ByVal nDestWidth As Long, _
	ByVal nDestHeight As Long, _
	ByVal hDIB As Long, _
	ByVal nXSrc As Long, _
	ByVal nYSrc As Long, _
	ByVal nSrcWidth As Long, _
	ByVal nSrcHeight As Long, _
	ByVal dwRop As Long, _
	ByVal iStretchMode As Long _
	) As Long
Public Declare Function DCtoDIB Lib "imgctl" ( _
	ByVal hdc As Long, _
	ByVal lXSrc As Long, _
	ByVal lYSrc As Long, _
	ByVal lWidth As Long, _
	ByVal lHeight As Long _
	) As Long	' HDIB
' for older than v1.08
Public Declare Function DIBDIBits Lib "imgctl" ( _
	ByVal hdc As Long, _
	ByVal nXDest As Long, _
	ByVal nYDest As Long, _
	ByVal nWidth As Long, _
	ByVal nHeight As Long, _
	ByVal hDIB As Long, _
	ByVal nXSrc As Long, _
	ByVal nYSrc As Long, _
	ByVal dwRop As Long _
	) As Long
Public Declare Function DIBStretchDIBits Lib "imgctl" ( _
	ByVal hdc As Long, _
	ByVal nXDest As Long, _
	ByVal nYDest As Long, _
	ByVal nDestWidth As Long, _
	ByVal nDestHeight As Long, _
	ByVal hDIB As Long, _
	ByVal nXSrc As Long, _
	ByVal nYSrc As Long, _
	ByVal nSrcWidth As Long, _
	ByVal nSrcHeight As Long, _
	ByVal dwRop As Long _
	) As Long
Public Declare Function DIBStretchDIBits2 Lib "imgctl" ( _
	ByVal hdc As Long, _
	ByVal nXDest As Long, _
	ByVal nYDest As Long, _
	ByVal nDestWidth As Long, _
	ByVal nDestHeight As Long, _
	ByVal hDIB As Long, _
	ByVal nXSrc As Long, _
	ByVal nYSrc As Long, _
	ByVal nSrcWidth As Long, _
	ByVal nSrcHeight As Long, _
	ByVal dwRop As Long, _
	ByVal iStretchMode As Long _
	) As Long

'---------------------------------------------------------------
' EOF