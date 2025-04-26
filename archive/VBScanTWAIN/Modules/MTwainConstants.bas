Attribute VB_Name = "MTwainConstants"
Option Explicit

'/****************************************************************************
'*   Generic     Constants   *
'****************************************************************************''/
Public Enum EGenericConstants
    TWON_ARRAY = 3
    TWON_ENUMERATION = 4
    TWON_ONEVALUE = 5
    TWON_RANGE = 6
    TWON_DSMCODEID = 63
    TWON_ICONID = 962
    TWON_DSMID = 461
    TWON_DONTCARE8 = &HFF&
    TWON_DONTCARE16 = &HFFFF&
    TWON_DONTCARE32 = &HFFFFFFFF
End Enum
                                                                        
Public Enum EMemoryFlags        'used   in  TW_MEMORY   structure,  *''/
    TWMF_APPOWNS = &H1
    TWMF_DSMOWNS = &H2
    TWMF_DSOWNS = &H4
    TWMF_POINTER = &H8
    TWMF_HANDLE = &H10
End Enum
                                                                        
Public Enum ETwainType
    TWTY_INT8 = &H0
    TWTY_INT16 = &H1
    TWTY_INT32 = &H2
        
    TWTY_UINT8 = &H3
    TWTY_UINT16 = &H4
    TWTY_UINT32 = &H5
        
    TWTY_BOOL = &H6
        
    TWTY_FIX32 = &H7
        
    TWTY_FRAME = &H8
        
    TWTY_STR32 = &H9
    TWTY_STR64 = &HA
    TWTY_STR128 = &HB
    TWTY_STR255 = &HC
    TWTY_HANDLE = &HF
End Enum
        
'/****************************************************************************
'*   Capability  =   Constants   *
'****************************************************************************''/
'/* CAP_ALARMS
Public Enum ECapAlarms
    TWAL_ALARM = 0
    TWAL_FEEDERERROR = 1
    TWAL_FEEDERWARNING = 2
    TWAL_BARCODE = 3
    TWAL_DOUBLEFEED = 4
    TWAL_JAM = 5
    TWAL_PATCHCODE = 6
    TWAL_POWER = 7
    TWAL_SKEW = 8
End Enum
                                                                        
'/* ICAP_AUTOSIZE
Public Enum ECapAutosize
    TWAS_NONE = 0
    TWAS_AUTO = 1
    TWAS_CURRENT = 2
End Enum
                                                                        
'/* TWEI_BARCODEROTATION
Public Enum ECapBarcodeRotation
    TWBCOR_ROT0 = 0
    TWBCOR_ROT90 = 1
    TWBCOR_ROT180 = 2
    TWBCOR_ROT270 = 3
    TWBCOR_ROTX = 4
End Enum
                                                                        
'/* ICAP_BARCODESEARCHMODE
Public Enum ECapBarcodeSearchmode
    TWBD_HORZ = 0
    TWBD_VERT = 1
    TWBD_HORZVERT = 2
    TWBD_VERTHORZ = 3
End Enum
                                                                        
'/* ICAP_BITORDER
Public Enum ECapBitorder
    TWBO_LSBFIRST = 0
    TWBO_MSBFIRST = 1
End Enum
                                                                        
'/* ICAP_AUTODISCARDBLANKPAGES
Public Enum ECapAutodiscardBlankpages
    TWBP_DISABLE = -2
    TWBP_AUTO = -1
End Enum
                                                                        
'/* ICAP_BITDEPTHREDUCTION
Public Enum ECapBitdepthReduction
    TWBR_THRESHOLD = 0
    TWBR_HALFTONE = 1
    TWBR_CUSTHALFTONE = 2
    TWBR_DIFFUSION = 3
    TWBR_DYNAMICTHRESHOLD = 4
End Enum
                                                                        
'/* ICAP_SUPPORTEDBARCODETYPES  =   and TWEI_BARCODETYPE    values*''/
Public Enum ESupportedBarcodeTypes
    TWBT_3OF9 = 0
    TWBT_2OF5INTERLEAVED = 1
    TWBT_2OF5NONINTERLEAVED = 2
    TWBT_CODE93 = 3
    TWBT_CODE128 = 4
    TWBT_UCC128 = 5
    TWBT_CODABAR = 6
    TWBT_UPCA = 7
    TWBT_UPCE = 8
    TWBT_EAN8 = 9
    TWBT_EAN13 = 10
    TWBT_POSTNET = 11
    TWBT_PDF417 = 12
    TWBT_2OF5INDUSTRIAL = 13
    TWBT_2OF5MATRIX = 14
    TWBT_2OF5DATALOGIC = 15
    TWBT_2OF5IATA = 16
    TWBT_3OF9FULLASCII = 17
    TWBT_CODABARWITHSTARTSTOP = 18
    TWBT_MAXICODE = 19
    TWBT_QRCODE = 20
End Enum
                                                                        
'/* ICAP_COMPRESSION    =   values*''/
Public Enum ECapCompression
    TWCP_NONE = 0
    TWCP_PACKBITS = 1
    TWCP_GROUP31D = 2
    TWCP_GROUP31DEOL = 3
    TWCP_GROUP32D = 4
    TWCP_GROUP4 = 5
    TWCP_JPEG = 6
    TWCP_LZW = 7
    TWCP_JBIG = 8
    TWCP_PNG = 9
    TWCP_RLE4 = 10
    TWCP_RLE8 = 11
    TWCP_BITFIELDS = 12
    TWCP_ZIP = 13
    TWCP_JPEG2000 = 14
End Enum
                                                                        
'/* CAP_CAMERASIDE  =   and TWEI_PAGESIDE   values  *''/
Public Enum ECapCameraSide
    TWCS_BOTH = 0
    TWCS_TOP = 1
    TWCS_BOTTOM = 2
End Enum
                                                                        
'/* CAP_DEVICEEVENT =   values  *''/
Public Enum ECapDeviceEvents
    TWDE_CUSTOMEVENTS = &H8000
    TWDE_CHECKAUTOMATICCAPTURE = 0
    TWDE_CHECKBATTERY = 1
    TWDE_CHECKDEVICEONLINE = 2
    TWDE_CHECKFLASH = 3
    TWDE_CHECKPOWERSUPPLY = 4
    TWDE_CHECKRESOLUTION = 5
    TWDE_DEVICEADDED = 6
    TWDE_DEVICEOFFLINE = 7
    TWDE_DEVICEREADY = 8
    TWDE_DEVICEREMOVED = 9
    TWDE_IMAGECAPTURED = 10
    TWDE_IMAGEDELETED = 11
    TWDE_PAPERDOUBLEFEED = 12
    TWDE_PAPERJAM = 13
    TWDE_LAMPFAILURE = 14
    TWDE_POWERSAVE = 15
    TWDE_POWERSAVENOTIFY = 16
End Enum
                                                                        
'/* TW_PASSTHRU,Direction   =   values, *''/
Public Enum EPassThru
    TWDR_GET = 1
    TWDR_SET = 2
End Enum
                                                                        
'/* TWEI_DESKEWSTATUS   =   values, *''/
Public Enum EDeskEWStatus
    TWDSK_SUCCESS = 0
    TWDSK_REPORTONLY = 1
    TWDSK_FAIL = 2
    TWDSK_DISABLED = 3
End Enum
                                                                        
'/* CAP_DUPLEX  =   values  *''/
Public Enum ECapDuplex
    TWDX_NONE = 0
    TWDX_1PASSDUPLEX = 1
    TWDX_2PASSDUPLEX = 2
End Enum

'/* CAP_FEEDERALIGNMENT =   values  *''/
Public Enum ECapFeederAlignment
    TWFA_NONE = 0
    TWFA_LEFT = 1
    TWFA_CENTER = 2
    TWFA_RIGHT = 3
End Enum

'/* ICAP_FEEDERTYPE =   values*''/
Public Enum ECapFeederType
    TWFE_GENERAL = 0
    TWFE_PHOTO = 1
End Enum

'/* ICAP_IMAGEFILEFORMAT    =   values  *''/
Public Enum ECapImageFileFormat
    TWFF_TIFF = 0
    TWFF_PICT = 1
    TWFF_BMP = 2
    TWFF_XBM = 3
    TWFF_JFIF = 4
    TWFF_FPX = 5
    TWFF_TIFFMULTI = 6
    TWFF_PNG = 7
    TWFF_SPIFF = 8
    TWFF_EXIF = 9
    TWFF_PDF = 10
    TWFF_JP2 = 11
    TWFF_JPX = 13
    TWFF_DEJAVU = 14
    TWFF_PDFA = 15
    TWFF_PDFA2 = 16
    TWFF_PDFRASTER = 17
End Enum

'/* ICAP_FLASHUSED2 =   values  *''/
Public Enum ECapFlashUsed2
    TWFL_NONE = 0
    TWFL_OFF = 1
    TWFL_ON = 2
    TWFL_AUTO = 3
    TWFL_REDEYE = 4
End Enum

'/* CAP_FEEDERORDER =   values  *''/
Public Enum ECapFeederOrder
    TWFO_FIRSTPAGEFIRST = 0
    TWFO_LASTPAGEFIRST = 1
End Enum

'/* CAP_FEEDERPOCKET    =   values*''/
Public Enum ECapFeederPocket
    TWFP_POCKETERROR = 0
    TWFP_POCKET1 = 1
    TWFP_POCKET2 = 2
    TWFP_POCKET3 = 3
    TWFP_POCKET4 = 4
    TWFP_POCKET5 = 5
    TWFP_POCKET6 = 6
    TWFP_POCKET7 = 7
    TWFP_POCKET8 = 8
    TWFP_POCKET9 = 9
    TWFP_POCKET10 = 10
    TWFP_POCKET11 = 11
    TWFP_POCKET12 = 12
    TWFP_POCKET13 = 13
    TWFP_POCKET14 = 14
    TWFP_POCKET15 = 15
    TWFP_POCKET16 = 16
End Enum

'/* ICAP_FLIPROTATION   =   values  *''/
Public Enum ECapFlipRotation
    TWFR_BOOK = 0
    TWFR_FANFOLD = 1
End Enum

'/* ICAP_FILTER =   values  *''/
Public Enum ECapFilter
    TWFT_RED = 0
    TWFT_GREEN = 1
    TWFT_BLUE = 2
    TWFT_NONE = 3
    TWFT_WHITE = 4
    TWFT_CYAN = 5
    TWFT_MAGENTA = 6
    TWFT_YELLOW = 7
    TWFT_BLACK = 8
End Enum

'/* TW_FILESYSTEM,FileType  =   values  *''/
Public Enum EFilesystemFileType
    TWFY_CAMERA = 0
    TWFY_CAMERATOP = 1
    TWFY_CAMERABOTTOM = 2
    TWFY_CAMERAPREVIEW = 3
    TWFY_DOMAIN = 4
    TWFY_HOST = 5
    TWFY_DIRECTORY = 6
    TWFY_IMAGE = 7
    TWFY_UNKNOWN = 8
End Enum

'/* ICAP_ICCPROFILE =   values  *''/
Public Enum ECapICCProfile
    TWIC_NONE = 0
    TWIC_LINK = 1
    TWIC_EMBED = 2
End Enum

'/* ICAP_IMAGEFILTER    =   values  *''/
Public Enum ECapImageFilter
    TWIF_NONE = 0
    TWIF_AUTO = 1
    TWIF_LOWPASS = 2
    TWIF_BANDPASS = 3
    TWIF_HIGHPASS = 4
    TWIF_TEXT = TWIF_BANDPASS
    TWIF_FINELINE = TWIF_HIGHPASS
End Enum

'/* ICAP_IMAGEMERGE =   values  *''/
Public Enum ECapImageMerge
    TWIM_NONE = 0
    TWIM_FRONTONTOP = 1
    TWIM_FRONTONBOTTOM = 2
    TWIM_FRONTONLEFT = 3
    TWIM_FRONTONRIGHT = 4
End Enum

'/* CAP_JOBCONTROL  =   values  *''/
Public Enum ECapJobControl
    TWJC_NONE = 0
    TWJC_JSIC = 1
    TWJC_JSIS = 2
    TWJC_JSXC = 3
    TWJC_JSXS = 4
End Enum

'/* ICAP_JPEGQUALITY    =   values  *''/
Public Enum ECapJpegQuality
    TWJQ_UNKNOWN = -4
    TWJQ_LOW = -3
    TWJQ_MEDIUM = -2
    TWJQ_HIGH = -1
End Enum

'/* ICAP_LIGHTPATH  =   values  *''/
Public Enum ECapLightPath
    TWLP_REFLECTIVE = 0
    TWLP_TRANSMISSIVE = 1
End Enum

'/* ICAP_LIGHTSOURCE    =   values  *''/
Public Enum ECapLightSource
    TWLS_RED = 0
    TWLS_GREEN = 1
    TWLS_BLUE = 2
    TWLS_NONE = 3
    TWLS_WHITE = 4
    TWLS_UV = 5
    TWLS_IR = 6
End Enum

'/* TWEI_MAGTYPE    =   values  *''/
Public Enum EMagType
    TWMD_MICR = 0
    TWMD_RAW = 1
    TWMD_INVALID = 2
End Enum

'/* ICAP_NOISEFILTER    =   values  *''/
Public Enum ECapNoiseFilter
    TWNF_NONE = 0
    TWNF_AUTO = 1
    TWNF_LONEPIXEL = 2
    TWNF_MAJORITYRULE = 3
End Enum

'/* ICAP_ORIENTATION    =   values  *''/
Public Enum ECapOrientation
    TWOR_ROT0 = 0
    TWOR_ROT90 = 1
    TWOR_ROT180 = 2
    TWOR_ROT270 = 3
    TWOR_PORTRAIT = TWOR_ROT0
    TWOR_LANDSCAPE = TWOR_ROT270
    TWOR_AUTO = 4
    TWOR_AUTOTEXT = 5
    TWOR_AUTOPICTURE = 6
End Enum

'/* ICAP_OVERSCAN   =   values  *''/
Public Enum ECapOverscan
    TWOV_NONE = 0
    TWOV_AUTO = 1
    TWOV_TOPBOTTOM = 2
    TWOV_LEFTRIGHT = 3
    TWOV_ALL = 4
End Enum

'/* Palette =   types   for TW_PALETTE8 *''/
Public Enum ECapPalette
    TWPA_RGB = 0
    TWPA_GRAY = 1
    TWPA_CMY = 2
End Enum

'/* ICAP_PLANARCHUNKY   =   values  *''/
Public Enum ECapPlanarChunky
    TWPC_CHUNKY = 0
    TWPC_PLANAR = 1
End Enum

'/* TWEI_PATCHCODE  =   values*''/
Public Enum EPatchcode
    TWPCH_PATCH1 = 0
    TWPCH_PATCH2 = 1
    TWPCH_PATCH3 = 2
    TWPCH_PATCH4 = 3
    TWPCH_PATCH6 = 4
    TWPCH_PATCHT = 5
End Enum

'/* ICAP_PIXELFLAVOR    =   values  *''/
Public Enum ECapPixelFlavor
    TWPF_CHOCOLATE = 0
    TWPF_VANILLA = 1
End Enum

'/* CAP_PRINTERMODE =   values  *''/
Public Enum ECapPrinterMode
    TWPM_SINGLESTRING = 0
    TWPM_MULTISTRING = 1
    TWPM_COMPOUNDSTRING = 2
End Enum

'/* CAP_PRINTER =   values  *''/
Public Enum ECapImprinter
    TWPR_IMPRINTERTOPBEFORE = 0
    TWPR_IMPRINTERTOPAFTER = 1
    TWPR_IMPRINTERBOTTOMBEFORE = 2
    TWPR_IMPRINTERBOTTOMAFTER = 3
End Enum

Public Enum ECapEndorser
    TWPR_ENDORSERTOPBEFORE = 4
    TWPR_ENDORSERTOPAFTER = 5
    TWPR_ENDORSERBOTTOMBEFORE = 6
    TWPR_ENDORSERBOTTOMAFTER = 7
End Enum

'/* CAP_PRINTERFONTSTYLE    =   Added   2,3 *''/
Public Enum ECapPrinterFontStyle
    TWPF_NORMAL = 0
    TWPF_BOLD = 1
    TWPF_ITALIC = 2
    TWPF_LARGESIZE = 3
    TWPF_SMALLSIZE = 4
End Enum

'/* CAP_PRINTERINDEXTRIGGER =   Added   2,3 *''/
Public Enum ECapPrinterIndexTrigger
    TWCT_PAGE = 0
    TWCT_PATCH1 = 1
    TWCT_PATCH2 = 2
    TWCT_PATCH3 = 3
    TWCT_PATCH4 = 4
    TWCT_PATCHT = 5
    TWCT_PATCH6 = 6
End Enum

'/* CAP_POWERSUPPLY =   values  *''/
Public Enum ECapPowerSupply
    TWPS_EXTERNAL = 0
    TWPS_BATTERY = 1
End Enum

'/* ICAP_PIXELTYPE  =   values  (PT_    means   Pixel   Type)   *''/
Public Enum ECapPixelType
    TWPT_BW = 0
    TWPT_GRAY = 1
    TWPT_RGB = 2
    TWPT_PALETTE = 3
    TWPT_CMY = 4
    TWPT_CMYK = 5
    TWPT_YUV = 6
    TWPT_YUVK = 7
    TWPT_CIEXYZ = 8
    TWPT_LAB = 9
    TWPT_SRGB = 10
    TWPT_SCRGB = 11
    TWPT_INFRARED = 16
End Enum

'/* CAP_SEGMENTED   =   values  *''/
Public Enum ECapSegmented
    TWSG_NONE = 0
    TWSG_AUTO = 1
    TWSG_MANUAL = 2
End Enum

'/* ICAP_FILMTYPE   =   values  *''/
Public Enum ECapFilmType
    TWFM_POSITIVE = 0
    TWFM_NEGATIVE = 1
End Enum

'/* CAP_DOUBLEFEEDDETECTION =   *''/
Public Enum ECapDoubleFeedDetection
    TWDF_ULTRASONIC = 0
    TWDF_BYLENGTH = 1
    TWDF_INFRARED = 2
End Enum

'/* CAP_DOUBLEFEEDDETECTIONSENSITIVITY  =   *''/
Public Enum ECapDoubleFeedDetectionSensibility
    TWUS_LOW = 0
    TWUS_MEDIUM = 1
    TWUS_HIGH = 2
End Enum

'/* CAP_DOUBLEFEEDDETECTIONRESPONSE =   *''/
Public Enum ECapDoubleFeedDetectionResponse
    TWDP_STOP = 0
    TWDP_STOPANDWAIT = 1
    TWDP_SOUND = 2
    TWDP_DONOTIMPRINT = 3
End Enum

'/* ICAP_MIRROR =   values  *''/
Public Enum ECapMirror
    TWMR_NONE = 0
    TWMR_VERTICAL = 1
    TWMR_HORIZONTAL = 2
End Enum
'/* ICAP_JPEGSUBSAMPLING    =   values  *''/
Public Enum ECapJpegSubsampling
    TWJS_444YCBCR = 0
    TWJS_444RGB = 1
    TWJS_422 = 2
    TWJS_421 = 3
    TWJS_411 = 4
    TWJS_420 = 5
    TWJS_410 = 6
    TWJS_311 = 7
End Enum

'/* CAP_PAPERHANDLING   =   values  *''/
Public Enum ECapPaperHandling
    TWPH_NORMAL = 0
    TWPH_FRAGILE = 1
    TWPH_THICK = 2
    TWPH_TRIFOLD = 3
    TWPH_PHOTOGRAPH = 4
End Enum

'/* CAP_INDICATORSMODE  =   values  *''/
Public Enum ECapIndicatorsMode
    TWCI_INFO = 0
    TWCI_WARNING = 1
    TWCI_ERROR = 2
    TWCI_WARMUP = 3
End Enum

'/* ICAP_SUPPORTEDSIZES =   values  (SS_    means   Supported   Sizes)  *''/
Public Enum ECapSupportedSizes
    TWSS_NONE = 0
    TWSS_A4 = 1
    TWSS_JISB5 = 2
    TWSS_USLETTER = 3
    TWSS_USLEGAL = 4
    TWSS_A5 = 5
    TWSS_ISOB4 = 6
    TWSS_ISOB6 = 7
    TWSS_USLEDGER = 9
    TWSS_USEXECUTIVE = 10
    TWSS_A3 = 11
    TWSS_ISOB3 = 12
    TWSS_A6 = 13
    TWSS_C4 = 14
    TWSS_C5 = 15
    TWSS_C6 = 16
    TWSS_4A0 = 17
    TWSS_2A0 = 18
    TWSS_A0 = 19
    TWSS_A1 = 20
    TWSS_A2 = 21
    TWSS_A7 = 22
    TWSS_A8 = 23
    TWSS_A9 = 24
    TWSS_A10 = 25
    TWSS_ISOB0 = 26
    TWSS_ISOB1 = 27
    TWSS_ISOB2 = 28
    TWSS_ISOB5 = 29
    TWSS_ISOB7 = 30
    TWSS_ISOB8 = 31
    TWSS_ISOB9 = 32
    TWSS_ISOB10 = 33
    TWSS_JISB0 = 34
    TWSS_JISB1 = 35
    TWSS_JISB2 = 36
    TWSS_JISB3 = 37
    TWSS_JISB4 = 38
    TWSS_JISB6 = 39
    TWSS_JISB7 = 40
    TWSS_JISB8 = 41
    TWSS_JISB9 = 42
    TWSS_JISB10 = 43
    TWSS_C0 = 44
    TWSS_C1 = 45
    TWSS_C2 = 46
    TWSS_C3 = 47
    TWSS_C7 = 48
    TWSS_C8 = 49
    TWSS_C9 = 50
    TWSS_C10 = 51
    TWSS_USSTATEMENT = 52
    TWSS_BUSINESSCARD = 53
    TWSS_MAXSIZE = 54
End Enum
'/* ICAP_XFERMECH   =   values  (SX_    means   Setup   XFer)   *''/
Public Enum ECapSetupXFer
    TWSX_NATIVE = 0
    TWSX_FILE = 1
    TWSX_MEMORY = 2
    TWSX_MEMFILE = 4
End Enum
'/* ICAP_UNITS  =   values  (UN_    means   UNits)  *''/
Public Enum ECapUnits
    TWUN_INCHES = 0
    TWUN_CENTIMETERS = 1
    TWUN_PICAS = 2
    TWUN_POINTS = 3
    TWUN_TWIPS = 4
    TWUN_PIXELS = 5
    TWUN_MILLIMETERS = 6
End Enum
        
'/****************************************************************************      =
'*   Country =   Constants   *
'****************************************************************************''/     =
Public Enum ETwainCountry
    TWCY_AFGHANISTAN = 1001
    TWCY_ALGERIA = 213
    TWCY_AMERICANSAMOA = 684
    TWCY_ANDORRA = 33
    TWCY_ANGOLA = 1002
    TWCY_ANGUILLA = 8090
    TWCY_ANTIGUA = 8091
    TWCY_ARGENTINA = 54
    TWCY_ARUBA = 297
    TWCY_ASCENSIONI = 247
    TWCY_AUSTRALIA = 61
    TWCY_AUSTRIA = 43
    TWCY_BAHAMAS = 8092
    TWCY_BAHRAIN = 973
    TWCY_BANGLADESH = 880
    TWCY_BARBADOS = 8093
    TWCY_BELGIUM = 32
    TWCY_BELIZE = 501
    TWCY_BENIN = 229
    TWCY_BERMUDA = 8094
    TWCY_BHUTAN = 1003
    TWCY_BOLIVIA = 591
    TWCY_BOTSWANA = 267
    TWCY_BRITAIN = 6
    TWCY_BRITVIRGINIS = 8095
    TWCY_BRAZIL = 55
    TWCY_BRUNEI = 673
    TWCY_BULGARIA = 359
    TWCY_BURKINAFASO = 1004
    TWCY_BURMA = 1005
    TWCY_BURUNDI = 1006
    TWCY_CAMAROON = 237
    TWCY_CANADA = 2
    TWCY_CAPEVERDEIS = 238
    TWCY_CAYMANIS = 8096
    TWCY_CENTRALAFREP = 1007
    TWCY_CHAD = 1008
    TWCY_CHILE = 56
    TWCY_CHINA = 86
    TWCY_CHRISTMASIS = 1009
    TWCY_COCOSIS = 1009
    TWCY_COLOMBIA = 57
    TWCY_COMOROS = 1010
    TWCY_CONGO = 1011
    TWCY_COOKIS = 1012
    TWCY_COSTARICA = 506
    TWCY_CUBA = 5
    TWCY_CYPRUS = 357
    TWCY_CZECHOSLOVAKIA = 42
    TWCY_DENMARK = 45
    TWCY_DJIBOUTI = 1013
    TWCY_DOMINICA = 8097
    TWCY_DOMINCANREP = 8098
    TWCY_EASTERIS = 1014
    TWCY_ECUADOR = 593
    TWCY_EGYPT = 20
    TWCY_ELSALVADOR = 503
    TWCY_EQGUINEA = 1015
    TWCY_ETHIOPIA = 251
    TWCY_FALKLANDIS = 1016
    TWCY_FAEROEIS = 298
    TWCY_FIJIISLANDS = 679
    TWCY_FINLAND = 358
    TWCY_FRANCE = 33
    TWCY_FRANTILLES = 596
    TWCY_FRGUIANA = 594
    TWCY_FRPOLYNEISA = 689
    TWCY_FUTANAIS = 1043
    TWCY_GABON = 241
    TWCY_GAMBIA = 220
    TWCY_GERMANY = 49
    TWCY_GHANA = 233
    TWCY_GIBRALTER = 350
    TWCY_GREECE = 30
    TWCY_GREENLAND = 299
    TWCY_GRENADA = 8099
    TWCY_GRENEDINES = 8015
    TWCY_GUADELOUPE = 590
    TWCY_GUAM = 671
    TWCY_GUANTANAMOBAY = 5399
    TWCY_GUATEMALA = 502
    TWCY_GUINEA = 224
    TWCY_GUINEABISSAU = 1017
    TWCY_GUYANA = 592
    TWCY_HAITI = 509
    TWCY_HONDURAS = 504
    TWCY_HONGKONG = 852
    TWCY_HUNGARY = 36
    TWCY_ICELAND = 354
    TWCY_INDIA = 91
    TWCY_INDONESIA = 62
    TWCY_IRAN = 98
    TWCY_IRAQ = 964
    TWCY_IRELAND = 353
    TWCY_ISRAEL = 972
    TWCY_ITALY = 39
    TWCY_IVORYCOAST = 225
    TWCY_JAMAICA = 8010
    TWCY_JAPAN = 81
    TWCY_JORDAN = 962
    TWCY_KENYA = 254
    TWCY_KIRIBATI = 1018
    TWCY_KOREA = 82
    TWCY_KUWAIT = 965
    TWCY_LAOS = 1019
    TWCY_LEBANON = 1020
    TWCY_LIBERIA = 231
    TWCY_LIBYA = 218
    TWCY_LIECHTENSTEIN = 41
    TWCY_LUXENBOURG = 352
    TWCY_MACAO = 853
    TWCY_MADAGASCAR = 1021
    TWCY_MALAWI = 265
    TWCY_MALAYSIA = 60
    TWCY_MALDIVES = 960
    TWCY_MALI = 1022
    TWCY_MALTA = 356
    TWCY_MARSHALLIS = 692
    TWCY_MAURITANIA = 1023
    TWCY_MAURITIUS = 230
    TWCY_MEXICO = 3
    TWCY_MICRONESIA = 691
    TWCY_MIQUELON = 508
    TWCY_MONACO = 33
    TWCY_MONGOLIA = 1024
    TWCY_MONTSERRAT = 8011
    TWCY_MOROCCO = 212
    TWCY_MOZAMBIQUE = 1025
    TWCY_NAMIBIA = 264
    TWCY_NAURU = 1026
    TWCY_NEPAL = 977
    TWCY_NETHERLANDS = 31
    TWCY_NETHANTILLES = 599
    TWCY_NEVIS = 8012
    TWCY_NEWCALEDONIA = 687
    TWCY_NEWZEALAND = 64
    TWCY_NICARAGUA = 505
    TWCY_NIGER = 227
    TWCY_NIGERIA = 234
    TWCY_NIUE = 1027
    TWCY_NORFOLKI = 1028
    TWCY_NORWAY = 47
    TWCY_OMAN = 968
    TWCY_PAKISTAN = 92
    TWCY_PALAU = 1029
    TWCY_PANAMA = 507
    TWCY_PARAGUAY = 595
    TWCY_PERU = 51
    TWCY_PHILLIPPINES = 63
    TWCY_PITCAIRNIS = 1030
    TWCY_PNEWGUINEA = 675
    TWCY_POLAND = 48
    TWCY_PORTUGAL = 351
    TWCY_QATAR = 974
    TWCY_REUNIONI = 1031
    TWCY_ROMANIA = 40
    TWCY_RWANDA = 250
    TWCY_SAIPAN = 670
    TWCY_SANMARINO = 39
    TWCY_SAOTOME = 1033
    TWCY_SAUDIARABIA = 966
    TWCY_SENEGAL = 221
    TWCY_SEYCHELLESIS = 1034
    TWCY_SIERRALEONE = 1035
    TWCY_SINGAPORE = 65
    TWCY_SOLOMONIS = 1036
    TWCY_SOMALI = 1037
    TWCY_SOUTHAFRICA = 27
    TWCY_SPAIN = 34
    TWCY_SRILANKA = 94
    TWCY_STHELENA = 1032
    TWCY_STKITTS = 8013
    TWCY_STLUCIA = 8014
    TWCY_STPIERRE = 508
    TWCY_STVINCENT = 8015
    TWCY_SUDAN = 1038
    TWCY_SURINAME = 597
    TWCY_SWAZILAND = 268
    TWCY_SWEDEN = 46
    TWCY_SWITZERLAND = 41
    TWCY_SYRIA = 1039
    TWCY_TAIWAN = 886
    TWCY_TANZANIA = 255
    TWCY_THAILAND = 66
    TWCY_TOBAGO = 8016
    TWCY_TOGO = 228
    TWCY_TONGAIS = 676
    TWCY_TRINIDAD = 8016
    TWCY_TUNISIA = 216
    TWCY_TURKEY = 90
    TWCY_TURKSCAICOS = 8017
    TWCY_TUVALU = 1040
    TWCY_UGANDA = 256
    TWCY_USSR = 7
    TWCY_UAEMIRATES = 971
    TWCY_UNITEDKINGDOM = 44
    TWCY_USA = 1
    TWCY_URUGUAY = 598
    TWCY_VANUATU = 1041
    TWCY_VATICANCITY = 39
    TWCY_VENEZUELA = 58
    TWCY_WAKE = 1042
    TWCY_WALLISIS = 1043
    TWCY_WESTERNSAHARA = 1044
    TWCY_WESTERNSAMOA = 1045
    TWCY_YEMEN = 1046
    TWCY_YUGOSLAVIA = 38
    TWCY_ZAIRE = 243
    TWCY_ZAMBIA = 260
    TWCY_ZIMBABWE = 263
    TWCY_ALBANIA = 355
    TWCY_ARMENIA = 374
    TWCY_AZERBAIJAN = 994
    TWCY_BELARUS = 375
    TWCY_BOSNIAHERZGO = 387
    TWCY_CAMBODIA = 855
    TWCY_CROATIA = 385
    TWCY_CZECHREPUBLIC = 420
    TWCY_DIEGOGARCIA = 246
    TWCY_ERITREA = 291
    TWCY_ESTONIA = 372
    TWCY_GEORGIA = 995
    TWCY_LATVIA = 371
    TWCY_LESOTHO = 266
    TWCY_LITHUANIA = 370
    TWCY_MACEDONIA = 389
    TWCY_MAYOTTEIS = 269
    TWCY_MOLDOVA = 373
    TWCY_MYANMAR = 95
    TWCY_NORTHKOREA = 850
    TWCY_PUERTORICO = 787
    TWCY_RUSSIA = 7
    TWCY_SERBIA = 381
    TWCY_SLOVAKIA = 421
    TWCY_SLOVENIA = 386
    TWCY_SOUTHKOREA = 82
    TWCY_UKRAINE = 380
    TWCY_USVIRGINIS = 340
    TWCY_VIETNAM = 84
End Enum
'/****************************************************************************      =
'*   Language    =   Constants   *
'****************************************************************************''/     =
Public Enum ETwainLanguage
    TWLG_USERLOCALE = -1
    TWLG_DAN = 0
    TWLG_DUT = 1
    TWLG_ENG = 2
    TWLG_FCF = 3
    TWLG_FIN = 4
    TWLG_FRN = 5
    TWLG_GER = 6
    TWLG_ICE = 7
    TWLG_ITN = 8
    TWLG_NOR = 9
    TWLG_POR = 10
    TWLG_SPA = 11
    TWLG_SWE = 12
    TWLG_USA = 13
    TWLG_AFRIKAANS = 14
    TWLG_ALBANIA = 15
    TWLG_ARABIC = 16
    TWLG_ARABIC_ALGERIA = 17
    TWLG_ARABIC_BAHRAIN = 18
    TWLG_ARABIC_EGYPT = 19
    TWLG_ARABIC_IRAQ = 20
    TWLG_ARABIC_JORDAN = 21
    TWLG_ARABIC_KUWAIT = 22
    TWLG_ARABIC_LEBANON = 23
    TWLG_ARABIC_LIBYA = 24
    TWLG_ARABIC_MOROCCO = 25
    TWLG_ARABIC_OMAN = 26
    TWLG_ARABIC_QATAR = 27
    TWLG_ARABIC_SAUDIARABIA = 28
    TWLG_ARABIC_SYRIA = 29
    TWLG_ARABIC_TUNISIA = 30
    TWLG_ARABIC_UAE = 31
    TWLG_ARABIC_YEMEN = 32
    TWLG_BASQUE = 33
    TWLG_BYELORUSSIAN = 34
    TWLG_BULGARIAN = 35
    TWLG_CATALAN = 36
    TWLG_CHINESE = 37
    TWLG_CHINESE_HONGKONG = 38
    TWLG_CHINESE_PRC = 39
    TWLG_CHINESE_SINGAPORE = 40
    TWLG_CHINESE_SIMPLIFIED = 41
    TWLG_CHINESE_TAIWAN = 42
    TWLG_CHINESE_TRADITIONAL = 43
    TWLG_CROATIA = 44
    TWLG_CZECH = 45
    TWLG_DANISH = TWLG_DAN
    TWLG_DUTCH = TWLG_DUT
    TWLG_DUTCH_BELGIAN = 46
    TWLG_ENGLISH = TWLG_ENG
    TWLG_ENGLISH_AUSTRALIAN = 47
    TWLG_ENGLISH_CANADIAN = 48
    TWLG_ENGLISH_IRELAND = 49
    TWLG_ENGLISH_NEWZEALAND = 50
    TWLG_ENGLISH_SOUTHAFRICA = 51
    TWLG_ENGLISH_UK = 52
    TWLG_ENGLISH_USA = TWLG_USA
    TWLG_ESTONIAN = 53
    TWLG_FAEROESE = 54
    TWLG_FARSI = 55
    TWLG_FINNISH = TWLG_FIN
    TWLG_FRENCH = TWLG_FRN
    TWLG_FRENCH_BELGIAN = 56
    TWLG_FRENCH_CANADIAN = TWLG_FCF
    TWLG_FRENCH_LUXEMBOURG = 57
    TWLG_FRENCH_SWISS = 58
    TWLG_GERMAN = TWLG_GER
    TWLG_GERMAN_AUSTRIAN = 59
    TWLG_GERMAN_LUXEMBOURG = 60
    TWLG_GERMAN_LIECHTENSTEIN = 61
    TWLG_GERMAN_SWISS = 62
    TWLG_GREEK = 63
    TWLG_HEBREW = 64
    TWLG_HUNGARIAN = 65
    TWLG_ICELANDIC = TWLG_ICE
    TWLG_INDONESIAN = 66
    TWLG_ITALIAN = TWLG_ITN
    TWLG_ITALIAN_SWISS = 67
    TWLG_JAPANESE = 68
    TWLG_KOREAN = 69
    TWLG_KOREAN_JOHAB = 70
    TWLG_LATVIAN = 71
    TWLG_LITHUANIAN = 72
    TWLG_NORWEGIAN = TWLG_NOR
    TWLG_NORWEGIAN_BOKMAL = 73
    TWLG_NORWEGIAN_NYNORSK = 74
    TWLG_POLISH = 75
    TWLG_PORTUGUESE = TWLG_POR
    TWLG_PORTUGUESE_BRAZIL = 76
    TWLG_ROMANIAN = 77
    TWLG_RUSSIAN = 78
    TWLG_SERBIAN_LATIN = 79
    TWLG_SLOVAK = 80
    TWLG_SLOVENIAN = 81
    TWLG_SPANISH = TWLG_SPA
    TWLG_SPANISH_MEXICAN = 82
    TWLG_SPANISH_MODERN = 83
    TWLG_SWEDISH = TWLG_SWE
    TWLG_THAI = 84
    TWLG_TURKISH = 85
    TWLG_UKRANIAN = 86
    TWLG_ASSAMESE = 87
    TWLG_BENGALI = 88
    TWLG_BIHARI = 89
    TWLG_BODO = 90
    TWLG_DOGRI = 91
    TWLG_GUJARATI = 92
    TWLG_HARYANVI = 93
    TWLG_HINDI = 94
    TWLG_KANNADA = 95
    TWLG_KASHMIRI = 96
    TWLG_MALAYALAM = 97
    TWLG_MARATHI = 98
    TWLG_MARWARI = 99
    TWLG_MEGHALAYAN = 100
    TWLG_MIZO = 101
    TWLG_NAGA = 102
    TWLG_ORISSI = 103
    TWLG_PUNJABI = 104
    TWLG_PUSHTU = 105
    TWLG_SERBIAN_CYRILLIC = 106
    TWLG_SIKKIMI = 107
    TWLG_SWEDISH_FINLAND = 108
    TWLG_TAMIL = 109
    TWLG_TELUGU = 110
    TWLG_TRIPURI = 111
    TWLG_URDU = 112
    TWLG_VIETNAMESE = 113
End Enum
        
'/****************************************************************************
'*   Data    =   Groups  *
'****************************************************************************''/
Public Enum EDataGroups
    DG_CONTROL = &H1&
    DG_IMAGE = &H2&
    DG_AUDIO = &H4&
'End Enum
'/* More    =   Data    Functionality   may be  added   in  the future,
'*   These   =   are for items   that    need    to  be  determined  before  DS  is  opened,
'*   NOTE:   =   Supported   Functionality   constants   must    be  powers  of  2   as  they    are
'*   used    =   as  bitflags    when    Application asks    DSM to  present a   list    of  DSs,
'*   to  =   support backward    capability  the App and DS  will    not use the fields
''/        =
    DF_DSM2 = &H10000000
    DF_APP2 = &H20000000
        
    DF_DS2 = &H40000000
        
    DG_MASK = &HFFFF&
        
'/****************************************************************************
'*   *
'****************************************************************************''/
    DAT_NULL = &H0
    DAT_CUSTOMBASE = &H8000
End Enum
'/* Data    =   Argument    Types   for the DG_CONTROL  Data    Group,  *''/
Public Enum EDataArgumentTypes
    DAT_CAPABILITY = &H1
    DAT_EVENT = &H2
    DAT_IDENTITY = &H3
    DAT_PARENT = &H4
    DAT_PENDINGXFERS = &H5
    DAT_SETUPMEMXFER = &H6
    DAT_SETUPFILEXFER = &H7
    DAT_STATUS = &H8
    DAT_USERINTERFACE = &H9
    DAT_XFERGROUP = &HA
    DAT_CUSTOMDSDATA = &HC
    DAT_DEVICEEVENT = &HD
    DAT_FILESYSTEM = &HE
    DAT_PASSTHRU = &HF
    DAT_CALLBACK = &H10
    DAT_STATUSUTF8 = &H11
    DAT_CALLBACK2 = &H12
    DAT_METRICS = &H13
    DAT_TWAINDIRECT = &H14
'End Enum
'/* Data    =   Argument    Types   for the DG_IMAGE    Data    Group,  *''/
                                                                        
    DAT_IMAGEINFO = &H101
    DAT_IMAGELAYOUT = &H102
    DAT_IMAGEMEMXFER = &H103
    DAT_IMAGENATIVEXFER = &H104
    DAT_IMAGEFILEXFER = &H105
    DAT_CIECOLOR = &H106
    DAT_GRAYRESPONSE = &H107
    DAT_RGBRESPONSE = &H108
    DAT_JPEGCOMPRESSION = &H109
    DAT_PALETTE8 = &H10A
    DAT_EXTIMAGEINFO = &H10B
    DAT_FILTER = &H10C
'End Enum
'/* Data    =   Argument    Types   for the DG_AUDIO    Data    Group,  *''/
                                                                        
    DAT_AUDIOFILEXFER = &H201
    DAT_AUDIOINFO = &H202
    DAT_AUDIONATIVEXFER = &H203
'End Enum
'/* misplaced   =   *''/
    DAT_ICCPROFILE = &H401
    DAT_IMAGEMEMFILEXFER = &H402
    DAT_ENTRYPOINT = &H403
End Enum
        
'/****************************************************************************      =
'*   Messages    =   *
'***************************************************************************''/     =
Public Enum ETwainMessages
'/* All =   message constants   are unique,
'*   Messages    =   are grouped according   to  which   DATs    they    are used    with,*''/
        
    MSG_NULL = &H0
    MSG_CUSTOMBASE = &H8000
        
'/* Generic =   messages    may be  used    with    any of  several DATs,   *''/
    MSG_GET = &H1
    MSG_GETCURRENT = &H2
    MSG_GETDEFAULT = &H3
    MSG_GETFIRST = &H4
    MSG_GETNEXT = &H5
    MSG_SET = &H6
    MSG_RESET = &H7
    MSG_QUERYSUPPORT = &H8
    MSG_GETHELP = &H9
    MSG_GETLABEL = &HA
    MSG_GETLABELENUM = &HB
    MSG_SETCONSTRAINT = &HC
        
'/* Messages    =   used    with    DAT_NULL    *''/
    MSG_XFERREADY = &H101
    MSG_CLOSEDSREQ = &H102
    MSG_CLOSEDSOK = &H103
    MSG_DEVICEEVENT = &H104
        
'/* Messages    =   used    with    a   pointer to  DAT_PARENT  data    *''/
    MSG_OPENDSM = &H301
    MSG_CLOSEDSM = &H302
        
'/* Messages    =   used    with    a   pointer to  a   DAT_IDENTITY    structure   *''/
    MSG_OPENDS = &H401
    MSG_CLOSEDS = &H402
    MSG_USERSELECT = &H403
        
'/* Messages    =   used    with    a   pointer to  a   DAT_USERINTERFACE   structure   *''/
    MSG_DISABLEDS = &H501
    MSG_ENABLEDS = &H502
    MSG_ENABLEDSUIONLY = &H503
        
'/* Messages    =   used    with    a   pointer to  a   DAT_EVENT   structure   *''/
    MSG_PROCESSEVENT = &H601
        
'/* Messages    =   used    with    a   pointer to  a   DAT_PENDINGXFERS    structure   *''/
    MSG_ENDXFER = &H701
    MSG_STOPFEEDER = &H702
        
'/* Messages    =   used    with    a   pointer to  a   DAT_FILESYSTEM  structure   *''/
    MSG_CHANGEDIRECTORY = &H801
    MSG_CREATEDIRECTORY = &H802
    MSG_DELETE = &H803
    MSG_FORMATMEDIA = &H804
    MSG_GETCLOSE = &H805
    MSG_GETFIRSTFILE = &H806
    MSG_GETINFO = &H807
    MSG_GETNEXTFILE = &H808
    MSG_RENAME = &H809
    MSG_COPY = &H80A
    MSG_AUTOMATICCAPTUREDIRECTORY = &H80B
        
'/* Messages    =   used    with    a   pointer to  a   DAT_PASSTHRU    structure   *''/
    MSG_PASSTHRU = &H901
        
'/* used    =   with    DAT_CALLBACK    *''/
    MSG_REGISTER_CALLBACK = &H902
        
'/* used    =   with    DAT_CAPABILITY  *''/
    MSG_RESETALL = &HA01
        
'/* used    =   with    DAT_TWAINDIRECT *''/
    MSG_SETTASK = &HB01
End Enum
''/****************************************************************************      =
'*   Capabilities    =   *
'****************************************************************************''/     =
Public Enum ECapabilities
    CAP_CUSTOMBASE = &H8000     '/* Base    of  custom  capabilities    *''/
'End Enum
'/* all =   data    sources are REQUIRED    to  support these   caps    *''/
    CAP_XFERCOUNT = &H1
'End Enum
'/* image   =   data    sources are REQUIRED    to  support these   caps    *''/
                                                                        
    ICAP_COMPRESSION = &H100
    ICAP_PIXELTYPE = &H101
    ICAP_UNITS = &H102
    ICAP_XFERMECH = &H103
'End Enum
'/* all =   data    sources MAY support these   caps    *''/
                                                                        
    CAP_AUTHOR = &H1000
    CAP_CAPTION = &H1001
    CAP_FEEDERENABLED = &H1002
    CAP_FEEDERLOADED = &H1003
    CAP_TIMEDATE = &H1004
    CAP_SUPPORTEDCAPS = &H1005
    CAP_EXTENDEDCAPS = &H1006
    CAP_AUTOFEED = &H1007
    CAP_CLEARPAGE = &H1008
    CAP_FEEDPAGE = &H1009
    CAP_REWINDPAGE = &H100A
    CAP_INDICATORS = &H100B
    CAP_PAPERDETECTABLE = &H100D
    CAP_UICONTROLLABLE = &H100E
    CAP_DEVICEONLINE = &H100F
    CAP_AUTOSCAN = &H1010
    CAP_THUMBNAILSENABLED = &H1011
    CAP_DUPLEX = &H1012
    CAP_DUPLEXENABLED = &H1013
    CAP_ENABLEDSUIONLY = &H1014
    CAP_CUSTOMDSDATA = &H1015
    CAP_ENDORSER = &H1016
    CAP_JOBCONTROL = &H1017
    CAP_ALARMS = &H1018
    CAP_ALARMVOLUME = &H1019
    CAP_AUTOMATICCAPTURE = &H101A
    CAP_TIMEBEFOREFIRSTCAPTURE = &H101B
    CAP_TIMEBETWEENCAPTURES = &H101C
    CAP_MAXBATCHBUFFERS = &H101E
    CAP_DEVICETIMEDATE = &H101F
    CAP_POWERSUPPLY = &H1020
    CAP_CAMERAPREVIEWUI = &H1021
    CAP_DEVICEEVENT = &H1022
    CAP_SERIALNUMBER = &H1024
    CAP_PRINTER = &H1026
    CAP_PRINTERENABLED = &H1027
    CAP_PRINTERINDEX = &H1028
    CAP_PRINTERMODE = &H1029
    CAP_PRINTERSTRING = &H102A
    CAP_PRINTERSUFFIX = &H102B
    CAP_LANGUAGE = &H102C
    CAP_FEEDERALIGNMENT = &H102D
    CAP_FEEDERORDER = &H102E
    CAP_REACQUIREALLOWED = &H1030
    CAP_BATTERYMINUTES = &H1032
    CAP_BATTERYPERCENTAGE = &H1033
    CAP_CAMERASIDE = &H1034
    CAP_SEGMENTED = &H1035
    CAP_CAMERAENABLED = &H1036
    CAP_CAMERAORDER = &H1037
    CAP_MICRENABLED = &H1038
    CAP_FEEDERPREP = &H1039
    CAP_FEEDERPOCKET = &H103A
    CAP_AUTOMATICSENSEMEDIUM = &H103B
    CAP_CUSTOMINTERFACEGUID = &H103C
    CAP_SUPPORTEDCAPSSEGMENTUNIQUE = &H103D
    CAP_SUPPORTEDDATS = &H103E
    CAP_DOUBLEFEEDDETECTION = &H103F
    CAP_DOUBLEFEEDDETECTIONLENGTH = &H1040
    CAP_DOUBLEFEEDDETECTIONSENSITIVITY = &H1041
    CAP_DOUBLEFEEDDETECTIONRESPONSE = &H1042
    CAP_PAPERHANDLING = &H1043
    CAP_INDICATORSMODE = &H1044
    CAP_PRINTERVERTICALOFFSET = &H1045
    CAP_POWERSAVETIME = &H1046
    CAP_PRINTERCHARROTATION = &H1047
    CAP_PRINTERFONTSTYLE = &H1048
    CAP_PRINTERINDEXLEADCHAR = &H1049
    CAP_PRINTERINDEXMAXVALUE = &H104A
    CAP_PRINTERINDEXNUMDIGITS = &H104B
    CAP_PRINTERINDEXSTEP = &H104C
    CAP_PRINTERINDEXTRIGGER = &H104D
    CAP_PRINTERSTRINGPREVIEW = &H104E
    CAP_SHEETCOUNT = &H104F
End Enum
                                                                        
'/* image       data    sources MAY support these   caps    *''/
Public Enum ECapabilityImage
    ICAP_AUTOBRIGHT = &H1100
    ICAP_BRIGHTNESS = &H1101
    ICAP_CONTRAST = &H1103
    ICAP_CUSTHALFTONE = &H1104
    ICAP_EXPOSURETIME = &H1105
    ICAP_FILTER = &H1106
    ICAP_FLASHUSED = &H1107
    ICAP_GAMMA = &H1108
    ICAP_HALFTONES = &H1109
    ICAP_HIGHLIGHT = &H110A
    ICAP_IMAGEFILEFORMAT = &H110C
    ICAP_LAMPSTATE = &H110D
    ICAP_LIGHTSOURCE = &H110E
    ICAP_ORIENTATION = &H1110
    ICAP_PHYSICALWIDTH = &H1111
    ICAP_PHYSICALHEIGHT = &H1112
    ICAP_SHADOW = &H1113
    ICAP_FRAMES = &H1114
    ICAP_XNATIVERESOLUTION = &H1116
    ICAP_YNATIVERESOLUTION = &H1117
    ICAP_XRESOLUTION = &H1118
    ICAP_YRESOLUTION = &H1119
    ICAP_MAXFRAMES = &H111A
    ICAP_TILES = &H111B
    ICAP_BITORDER = &H111C
    ICAP_CCITTKFACTOR = &H111D
    ICAP_LIGHTPATH = &H111E
    ICAP_PIXELFLAVOR = &H111F
    ICAP_PLANARCHUNKY = &H1120
    ICAP_ROTATION = &H1121
    ICAP_SUPPORTEDSIZES = &H1122
    ICAP_THRESHOLD = &H1123
    ICAP_XSCALING = &H1124
    ICAP_YSCALING = &H1125
    ICAP_BITORDERCODES = &H1126
    ICAP_PIXELFLAVORCODES = &H1127
    ICAP_JPEGPIXELTYPE = &H1128
    ICAP_TIMEFILL = &H112A
    ICAP_BITDEPTH = &H112B
    ICAP_BITDEPTHREDUCTION = &H112C
    ICAP_UNDEFINEDIMAGESIZE = &H112D
    ICAP_IMAGEDATASET = &H112E
    ICAP_EXTIMAGEINFO = &H112F
    ICAP_MINIMUMHEIGHT = &H1130
    ICAP_MINIMUMWIDTH = &H1131
    ICAP_AUTODISCARDBLANKPAGES = &H1134
    ICAP_FLIPROTATION = &H1136
    ICAP_BARCODEDETECTIONENABLED = &H1137
    ICAP_SUPPORTEDBARCODETYPES = &H1138
    ICAP_BARCODEMAXSEARCHPRIORITIES = &H1139
    ICAP_BARCODESEARCHPRIORITIES = &H113A
    ICAP_BARCODESEARCHMODE = &H113B
    ICAP_BARCODEMAXRETRIES = &H113C
    ICAP_BARCODETIMEOUT = &H113D
    ICAP_ZOOMFACTOR = &H113E
    ICAP_PATCHCODEDETECTIONENABLED = &H113F
    ICAP_SUPPORTEDPATCHCODETYPES = &H1140
    ICAP_PATCHCODEMAXSEARCHPRIORITIES = &H1141
    ICAP_PATCHCODESEARCHPRIORITIES = &H1142
    ICAP_PATCHCODESEARCHMODE = &H1143
    ICAP_PATCHCODEMAXRETRIES = &H1144
    ICAP_PATCHCODETIMEOUT = &H1145
    ICAP_FLASHUSED2 = &H1146
    ICAP_IMAGEFILTER = &H1147
    ICAP_NOISEFILTER = &H1148
    ICAP_OVERSCAN = &H1149
    ICAP_AUTOMATICBORDERDETECTION = &H1150
    ICAP_AUTOMATICDESKEW = &H1151
    ICAP_AUTOMATICROTATE = &H1152
    ICAP_JPEGQUALITY = &H1153
    ICAP_FEEDERTYPE = &H1154
    ICAP_ICCPROFILE = &H1155
    ICAP_AUTOSIZE = &H1156
    ICAP_AUTOMATICCROPUSESFRAME = &H1157
    ICAP_AUTOMATICLENGTHDETECTION = &H1158
    ICAP_AUTOMATICCOLORENABLED = &H1159
    ICAP_AUTOMATICCOLORNONCOLORPIXELTYPE = &H115A
    ICAP_COLORMANAGEMENTENABLED = &H115B
    ICAP_IMAGEMERGE = &H115C
    ICAP_IMAGEMERGEHEIGHTTHRESHOLD = &H115D
    ICAP_SUPPORTEDEXTIMAGEINFO = &H115E
    ICAP_FILMTYPE = &H115F
    ICAP_MIRROR = &H1160
    ICAP_JPEGSUBSAMPLING = &H1161
'End Enum
'Public Enum
'/* image   =   data    sources MAY support these   audio   caps    *''/
    ACAP_XFERMECH = &H1202
End Enum
        
'/***************************************************************************       =
'*   Extended    =   Image   Info    Attributes  section Added   1,7 *
'***************************************************************************''/      =
Public Enum EExtendedImageInfo
    TWEI_BARCODEX = &H1200
    TWEI_BARCODEY = &H1201
    TWEI_BARCODETEXT = &H1202
    TWEI_BARCODETYPE = &H1203
    TWEI_DESHADETOP = &H1204
    TWEI_DESHADELEFT = &H1205
    TWEI_DESHADEHEIGHT = &H1206
    TWEI_DESHADEWIDTH = &H1207
    TWEI_DESHADESIZE = &H1208
    TWEI_SPECKLESREMOVED = &H1209
    TWEI_HORZLINEXCOORD = &H120A
    TWEI_HORZLINEYCOORD = &H120B
    TWEI_HORZLINELENGTH = &H120C
    TWEI_HORZLINETHICKNESS = &H120D
    TWEI_VERTLINEXCOORD = &H120E
    TWEI_VERTLINEYCOORD = &H120F
    TWEI_VERTLINELENGTH = &H1210
    TWEI_VERTLINETHICKNESS = &H1211
    TWEI_PATCHCODE = &H1212
    TWEI_ENDORSEDTEXT = &H1213
    TWEI_FORMCONFIDENCE = &H1214
    TWEI_FORMTEMPLATEMATCH = &H1215
    TWEI_FORMTEMPLATEPAGEMATCH = &H1216
    TWEI_FORMHORZDOCOFFSET = &H1217
    TWEI_FORMVERTDOCOFFSET = &H1218
    TWEI_BARCODECOUNT = &H1219
    TWEI_BARCODECONFIDENCE = &H121A
    TWEI_BARCODEROTATION = &H121B
    TWEI_BARCODETEXTLENGTH = &H121C
    TWEI_DESHADECOUNT = &H121D
    TWEI_DESHADEBLACKCOUNTOLD = &H121E
    TWEI_DESHADEBLACKCOUNTNEW = &H121F
    TWEI_DESHADEBLACKRLMIN = &H1220
    TWEI_DESHADEBLACKRLMAX = &H1221
    TWEI_DESHADEWHITECOUNTOLD = &H1222
    TWEI_DESHADEWHITECOUNTNEW = &H1223
    TWEI_DESHADEWHITERLMIN = &H1224
    TWEI_DESHADEWHITERLAVE = &H1225
    TWEI_DESHADEWHITERLMAX = &H1226
    TWEI_BLACKSPECKLESREMOVED = &H1227
    TWEI_WHITESPECKLESREMOVED = &H1228
    TWEI_HORZLINECOUNT = &H1229
    TWEI_VERTLINECOUNT = &H122A
    TWEI_DESKEWSTATUS = &H122B
    TWEI_SKEWORIGINALANGLE = &H122C
    TWEI_SKEWFINALANGLE = &H122D
    TWEI_SKEWCONFIDENCE = &H122E
    TWEI_SKEWWINDOWX1 = &H122F
    TWEI_SKEWWINDOWY1 = &H1230
    TWEI_SKEWWINDOWX2 = &H1231
    TWEI_SKEWWINDOWY2 = &H1232
    TWEI_SKEWWINDOWX3 = &H1233
    TWEI_SKEWWINDOWY3 = &H1234
    TWEI_SKEWWINDOWX4 = &H1235
    TWEI_SKEWWINDOWY4 = &H1236
    TWEI_BOOKNAME = &H1238
    TWEI_CHAPTERNUMBER = &H1239
    TWEI_DOCUMENTNUMBER = &H123A
    TWEI_PAGENUMBER = &H123B
    TWEI_CAMERA = &H123C
    TWEI_FRAMENUMBER = &H123D
    TWEI_FRAME = &H123E
    TWEI_PIXELFLAVOR = &H123F
    TWEI_ICCPROFILE = &H1240
    TWEI_LASTSEGMENT = &H1241
    TWEI_SEGMENTNUMBER = &H1242
    TWEI_MAGDATA = &H1243
    TWEI_MAGTYPE = &H1244
    TWEI_PAGESIDE = &H1245
    TWEI_FILESYSTEMSOURCE = &H1246
    TWEI_IMAGEMERGED = &H1247
    TWEI_MAGDATALENGTH = &H1248
    TWEI_PAPERCOUNT = &H1249
    TWEI_PRINTERTEXT = &H124A
    TWEI_TWAINDIRECTMETADATA = &H124B
End Enum
Public Enum EExtendedJpeg
    TWEJ_NONE = &H0
    TWEJ_MIDSEPARATOR = &H1
    TWEJ_PATCH1 = &H2
    TWEJ_PATCH2 = &H3
    TWEJ_PATCH3 = &H4
    TWEJ_PATCH4 = &H5
    TWEJ_PATCH6 = &H6
    TWEJ_PATCHT = &H7
End Enum
        
'/***************************************************************************       =
'*   Return  =   Codes   and Condition   Codes   section *
'***************************************************************************''/      =
Public Enum EReturnCodes
    TWRC_CUSTOMBASE = &H8000
    
    TWRC_SUCCESS = 0
    TWRC_FAILURE = 1
    TWRC_CHECKSTATUS = 2
    TWRC_CANCEL = 3
    TWRC_DSEVENT = 4
    TWRC_NOTDSEVENT = 5
    TWRC_XFERDONE = 6
    TWRC_ENDOFLIST = 7
    TWRC_INFONOTSUPPORTED = 8
    TWRC_DATANOTAVAILABLE = 9
    TWRC_BUSY = 10
    TWRC_SCANNERLOCKED = 11
End Enum
'/* Condition   =   Codes:  Application gets    these   by  doing   DG_CONTROL  DAT_STATUS  MSG_GET,    *''/
Public Enum EConditionCodes
    TWCC_CUSTOMBASE = &H8000&
    '
    TWCC_SUCCESS = 0
    TWCC_BUMMER = 1
    TWCC_LOWMEMORY = 2
    TWCC_NODS = 3
    TWCC_MAXCONNECTIONS = 4
    TWCC_OPERATIONERROR = 5
    TWCC_BADCAP = 6
    '
    '
    TWCC_BADPROTOCOL = 9
    TWCC_BADVALUE = 10
    TWCC_SEQERROR = 11
    TWCC_BADDEST = 12
    TWCC_CAPUNSUPPORTED = 13
    TWCC_CAPBADOPERATION = 14
    TWCC_CAPSEQERROR = 15
    TWCC_DENIED = 16
    TWCC_FILEEXISTS = 17
    TWCC_FILENOTFOUND = 18
    TWCC_NOTEMPTY = 19
    TWCC_PAPERJAM = 20
    TWCC_PAPERDOUBLEFEED = 21
    TWCC_FILEWRITEERROR = 22
    TWCC_CHECKDEVICEONLINE = 23
    TWCC_INTERLOCK = 24
    TWCC_DAMAGEDCORNER = 25
    TWCC_FOCUSERROR = 26
    TWCC_DOCTOOLIGHT = 27
    TWCC_DOCTOODARK = 28
    TWCC_NOMEDIA = 29
End Enum
'/* bit =   patterns:   for query   the operation   that    are supported   by  the data    source  on  a   capability  *''/
'/* Application =   gets    these   through DG_CONTROL''/DAT_CAPABILITY''/MSG_QUERYSUPPORT  *''/
Public Enum EQualityControl
    TWQC_GET = &H1
    TWQC_SET = &H2
    TWQC_GETDEFAULT = &H4
    TWQC_GETCURRENT = &H8
    TWQC_RESET = &H10
    TWQC_SETCONSTRAINT = &H20
    TWQC_GETHELP = &H100
    TWQC_GETLABEL = &H200
    TWQC_GETLABELENUM = &H400
End Enum
'        =
''/****************************************************************************      =
'*   Deprecated  =   Items   *
'****************************************************************************''/     =
'#if defined(WIN32)  =   ||  defined(WIN64)
'    TW_HUGE =
'#elif   !defined(TWH_CMP_GNU)   =
'    TW_HUGE = huge
'#else       =
'    TW_HUGE =
'#endif      =
'        =
''typedef    BYTE    =   TW_HUGE *   HPBYTE;
''typedef    void    =   TW_HUGE *   HPVOID;
'        =
''typedef    unsigned    =   char    TW_STR1024[1026],   FAR *pTW_STR1026,   FAR *pTW_STR1024;
''typedef    wchar_t =   TW_UNI512[512], FAR *pTW_UNI512;
'        =
'    TWTY_STR1024 = &HD
'    TWTY_UNI512 = &HE
'        =
'    TWFF_JPN = 12
'        =
'    DAT_TWUNKIDENTITY = &HB
'    DAT_SETUPFILEXFER2 = &H301
'        =
'    CAP_CLEARBUFFERS = &H101D
'    CAP_SUPPORTEDCAPSEXT = &H100C
'    CAP_FILESYSTEM  =   '/''/&H????
'    CAP_PAGEMULTIPLEACQUIRE = &H1023
'    CAP_PAPERBINDING = &H102F
'    CAP_PASSTHRU = &H1031
'    CAP_POWERDOWNTIME = &H1034
'    ACAP_AUDIOFILEFORMAT = &H1201
'        =
'    MSG_CHECKSTATUS = &H201
'        =
'    MSG_INVOKE_CALLBACK = &H903     '/* Mac Only,   deprecated  -   use DAT_NULL    and MSG_xxx instead *''/
'        =
'    TWQC_CONSTRAINABLE = &H40
'        =
'    TWSX_FILE2 = 3
'        =
''/* CAP_FILESYSTEM  =   values  (FS_    means   file    system) *''/
'    TWFS_FILESYSTEM = 0
'    TWFS_RECURSIVEDELETE = 1
'        =
''/* ICAP_PIXELTYPE  =   values  (PT_    means   Pixel   Type)   *''/
'
'    TWPT_SRGB64 = 11
'    TWPT_BGR = 12
'    TWPT_CIELAB = 13
'    TWPT_CIELUV = 14
'    TWPT_YCBCR = 15
'
'        =
''/* ICAP_SUPPORTEDSIZES =   values  (SS_    means   Supported   Sizes)  *''/
'
'    TWSS_B = 8
'    TWSS_A4LETTER = TWSS_A4
'    TWSS_B3 = TWSS_ISOB3
'    TWSS_B4 = TWSS_ISOB4
'    TWSS_B6 = TWSS_ISOB6
'    TWSS_B5LETTER = TWSS_JISB5
'
'        =
''/* ACAP_AUDIOFILEFORMAT    =   values  (AF_    means   audio   format),    Added   1,8 *''/
'
'    TWAF_WAV = 0
'    TWAF_AIFF = 1
'    TWAF_AU = 3
'    TWAF_SND = 4
'
'        =
''/* CAP_CLEARBUFFERS    =   values  *''/
'
'    TWCB_AUTO = 0
'    TWCB_CLEAR = 1
'    TWCB_NOCLEAR = 2
'
