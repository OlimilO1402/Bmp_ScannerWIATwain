Attribute VB_Name = "MTwainTypes"
Option Explicit
'/****************************************************************************
'*  Type    Definitions *
'****************************************************************************''/

'/* Fixed   point   Structure type    *''/
Public Type TW_FIX32
    Whole            As Integer
    Frac             As Integer
End Type

'/* Defines a   frame   rectangle   in  ICAP_UNITS  coordinates *''/
Public Type TW_FRAME
    Left             As TW_FIX32
    Top              As TW_FIX32
    Right            As TW_FIX32
    Bottom           As TW_FIX32
End Type

'/* Defines the parameters  used    for channel-specific    transformation  *''/
Public Type TW_DECODEFUNCTION
    StartIn          As TW_FIX32
    BreakIn          As TW_FIX32
    EndIn            As TW_FIX32
    StartOut         As TW_FIX32
    BreakOut         As TW_FIX32
    EndOut           As TW_FIX32
    Gamma            As TW_FIX32
    SampleCount      As TW_FIX32
End Type

'/* Stores  a   Fixed   point   number  in  two parts   a   whole   and a   fractional  part    *''/
Public Type TW_TRANSFORMSTAGE
    Decode(3)           As TW_DECODEFUNCTION
    Mix(0 To 2, 0 To 2) As TW_FIX32
End Type

'/* Container   for array   of  values  *''/
Public Type TW_ARRAY
    ItemType         As Integer
    NumItems         As Long
    ItemList(0 To 1) As Byte
End Type

'/* Information about   audio   data    *''/
Public Type TW_AUDIOINFO
    NName(0 To 255)  As Byte
    Reserved         As Long
End Type

'/* Used    to  register    callbacks   *''/
Public Type TW_CALLBACK
    CallBackProc    As LongPtr 'TW_MEMREF
'#if    defined(__APPLE__)  '/* cf: Mac version of  TWAINh  *''/
    RefCon          As LongPtr 'TW_MEMREF
'#else
    'RefCon  As LongPtr
'#endif
    Message         As Integer
End Type

'/* Used    to  register    callbacks   *''/
Public Type TW_CALLBACK2
    CallBackProc    As LongPtr 'TW_MEMREF
    RefCon          As LongPtr
    Message         As Integer
End Type

'/* Used    by  application to  get''/set   capability  from''/in   a   data    source  *''/
Public Type TW_CAPABILITY
    Cap             As Integer
    ConType         As Integer
    hContainer      As LongPtr
End Type

'/* Defines a   CIE XYZ space   tri-stimulus    value   *''/
Public Type TW_CIEPOINT
    X   As TW_FIX32
    Y   As TW_FIX32
    Z   As TW_FIX32
End Type

'/* Defines the mapping from    an  RGB color   space   device  into    CIE 1931    (XYZ)   color   space   *''/
Public Type TW_CIECOLOR
    ColorSpace      As Integer
    LowEndian       As Integer
    DeviceDependent As Integer
    VersionNumber   As Long
    StageABC        As TW_TRANSFORMSTAGE
    StageLMN        As TW_TRANSFORMSTAGE
    WhitePoint      As TW_CIEPOINT
    BlackPoint      As TW_CIEPOINT
    WhitePaper      As TW_CIEPOINT
    BlackInk        As TW_CIEPOINT
    Samples(1)      As TW_FIX32
End Type

'/* Allows  for a   data    source  and application to  pass    custom  data    to  each    other   *''/
Public Type TW_CUSTOMDSDATA
    InfoLength      As Long
    hData           As LongPtr
End Type

'/* Provides    information about   the Event   that    was raised  by  the Source  *''/
Public Type TW_DEVICEEVENT
    Event                  As Long
    DeviceName(0 To 255)   As Byte
    BatteryMinutes         As Long
    BatteryPercentage      As Integer
    PowerSupply            As Long
    XResolution            As TW_FIX32
    YResolution            As TW_FIX32
    FlashUsed2             As Long
    AutomaticCapture       As Long
    TimeBeforeFirstCapture As Long
    TimeBetweenCaptures    As Long
End Type

'/* This    Structure holds   the tri-stimulus    color   palette information for TW_PALETTE8 Structures*''/
Public Type TW_ELEMENT8
    Index       As Byte
    Channel1    As Byte
    Channel2    As Byte
    Channel3    As Byte
End Type

'/* Stores  a   group   of  individual  values  describing  a   capability  *''/
Public Type TW_ENUMERATION
    ItemType        As Integer
    NumItems        As Long
    CurrentIndex    As Long
    DefaultIndex    As Long
    ItemList(1)     As Byte
End Type

'/* Used    to  pass    application events''/messages   from    the application to  the Source  *''/
Public Type TW_EVENT
    pEvent          As LongPtr 'TW_MEMREF
    TWMessage       As Integer
End Type

'/* This    Structure is  used    to  pass    specific    information between the data    source  and the application *''/
Public Type TW_INFO
    InfoID          As Integer
    ItemType        As Integer
    NumItems        As Integer
    'Union
    UnionReturnOrCondCode As Integer
    'CondCode        As Integer      '/''/   Deprecated  do  not use
    'End Type
    Item            As LongPtr
End Type

Public Type TW_EXTIMAGEINFO
    NumInfos    As Long
    Info(1)     As TW_INFO
End Type

'/* Provides    information about   the currently   selected    device  *''/
Public Type TW_FILESYSTEM
    InputName(0 To 255)  As Byte
    OutputName(0 To 255) As Byte
    Context              As LongPtr 'TW_MEMREF
    '#Union
        UnionRecursiveOrSubdirectories As Long
        'Subdirectories   As Boolean 'TW_BOOL
    'End Type
    'Union
        UnionFileTypeOrFileSystemType As Long
        'FileSystemType   As Long
    'End Type
    Size                      As Long
    CreateTimeDate(0 To 33)   As Byte
    ModifiedTimeDate(0 To 33) As Byte
    FreeSpace                 As Long
    NewImageSize              As Long
    NumberOfFiles             As Long
    NumberOfSnippets          As Long
    DeviceGroupMask           As Long
    Reserved(508)             As Byte 'TW_INT8
End Type

'/* This    Structure is  used    by  the application to  specify a   set of  mapping values  to  be  applied to  grayscale   data    *''/
Public Type TW_GRAYRESPONSE
    Response(1) As TW_ELEMENT8
End Type

'/* A   general way to  describe    the version of  software    that    is  running *''/
Public Type TW_VERSION
    MajorNum      As Integer
    MinorNum      As Integer
    Language      As Integer
    Country       As Integer
    Info(0 To 33) As Byte
End Type

'/* Provides    identification  information about   a   TWAIN   entity*''/
Public Type TW_IDENTITY
'#if    defined(__APPLE__)  '/* cf: Mac version of  TWAINh  *''/
    Id  As LongPtr 'TW_MEMREF
'#else
    'Id  As Long
'#endif
    Version                As TW_VERSION
    ProtocolMajor          As Integer
    ProtocolMinor          As Integer
    SupportedGroups        As Long
    Manufacturer(0 To 33)  As Byte
    ProductFamily(0 To 33) As Byte
    ProductName(0 To 33)   As Byte
End Type

'/* Describes   the real    image   data    that    is  the complete    image   being   transferred between the Source  and application *''/
Public Type TW_IMAGEINFO
    XResolution      As TW_FIX32
    YResolution      As TW_FIX32
    ImageWidth       As Long
    ImageLength      As Long
    SamplesPerPixel  As Integer
    BitsPerSample(8) As Integer
    BitsPerPixel     As Integer
    Planar           As Boolean 'TW_BOOL
    PixelType        As Integer
    Compression      As Integer
End Type

'/* Involves    information about   the original    size    of  the acquired    image   *''/
Public Type TW_IMAGELAYOUT
    Frame          As TW_FRAME
    DocumentNumber As Long
    PageNumber     As Long
    FrameNumber    As Long
End Type

'/* Provides    information for managing    memory  buffers *''/
Public Type TW_MEMORY
    Flags   As Long
    Length  As Long
    TheMem  As LongPtr 'TW_MEMREF
End Type

'/* Describes   the form    of  the acquired    data    being   passed  from    the Source  to  the application*''/
Public Type TW_IMAGEMEMXFER
    Compression  As Integer
    BytesPerRow  As Long
    Columns      As Long
    Rows         As Long
    XOffset      As Long
    YOffset      As Long
    BytesWritten As Long
    Memory       As TW_MEMORY
End Type

'/* Describes   the information necessary   to  transfer    a   JPEG-compressed image   *''/
Public Type TW_JPEGCOMPRESSION
    ColorSpace       As Integer
    SubSampling      As Long
    NumComponents    As Integer
    RestartFrequency As Integer
    QuantMap(4)      As Integer
    QuantTable(4)    As TW_MEMORY
    HuffmanMap(4)    As Integer
    HuffmanDC(2)     As TW_MEMORY
    HuffmanAC(2)     As TW_MEMORY
End Type

'/* Collects    scanning    metrics after   returning   to  state   4   *''/
Public Type TW_METRICS
    SizeOf           As Long
    ImageCount       As Long
    SheetCount       As Long
End Type

'/* Stores  a   single  value   (item)  which   describes   a   capability  *''/
Public Type TW_ONEVALUE
    ItemType         As Integer
    Item             As Long
End Type

'/* This    Structure holds   the color   palette information *''/
Public Type TW_PALETTE8
    NumColors        As Integer
    PaletteType      As Integer
    Colors(256)      As TW_ELEMENT8
End Type

'/* Used    to  bypass  the TWAIN   protocol    when    communicating   with    a   device  *''/
Public Type TW_PASSTHRU
    pCommand    As LongPtr 'TW_MEMREF
    CommandBytes    As Long
    Direction   As Long
    pData   As LongPtr 'TW_MEMREF
    DataBytes   As Long
    DataBytesXfered As Long
End Type

'/* This    Structure tells   the application how many    more    complete    transfers   the Source  currently   has available   *''/
Public Type TW_PENDINGXFERS
    Count   As Integer
    'Union
    'EOJ As Long
    'Reserved    As Long
'#if    defined(__APPLE__)  '/* cf: Mac version of  TWAINh  *''/
    'Union TW_JOBCONTROL
    EOJ As Long
    'Reserved    As Long
    'End Type
    '#endif
    'End Type
End Type
                                                                                        
'/* Stores  a   range   of  individual  values  describing  a   capability  *''/
Public Type TW_RANGE
    ItemType    As Integer
    MinValue    As Long
    MaxValue    As Long
    StepSize    As Long
    DefaultValue    As Long
    CurrentValue    As Long
End Type
                                                                                        
'/* This    Structure is  used    by  the application to  specify a   set of  mapping values  to  be  applied to  RGB color   data    *''/
Public Type TW_RGBRESPONSE
    Response(1) As TW_ELEMENT8
End Type

'/* Describes   the file    format  and file    specification   information for a   transfer    through a   disk    file    *''/
Public Type TW_SETUPFILEXFER
    FileName(0 To 255) As Byte
    Format  As Integer
    VRefNum As Integer
End Type

'/* Provides    the application information about   the Source's    requirements    and preferences regarding   allocation  of  transfer    buffer(s)   *''/
Public Type TW_SETUPMEMXFER
    MinBufSize  As Long
    MaxBufSize  As Long
    Preferred   As Long
End Type

'/* Describes   the status  of  a   source  *''/
Public Type TW_STATUS
    ConditionCode   As Integer
    'Union
    Data    As Integer
    'Reserved    As Integer
    'End Type
End Type

'/* Translates  the contents    of  Status  into    a   localized   UTF8string  *''/
Public Type TW_STATUSUTF8
    Status               As TW_STATUS
    Size                 As Long
    UTF8string           As LongPtr
End Type

Public Type TW_TWAINDIRECT
    SizeOf               As Long
    CommunicationManager As Integer
    Send                 As LongPtr
    SendSize             As Long
    Receive              As LongPtr
    ReceiveSize          As Long
End Type

'/* This    Structure is  used    to  handle  the user    interface   coordination    between an  application and a   Source  *''/
Public Type TW_USERINTERFACE
    ShowUI               As Boolean 'TW_BOOL
    ModalUI              As Boolean 'TW_BOOL
    hParent              As LongPtr
End Type

'Ordinal: 4(0x0004); Hint: 0(0x0000); Entry-Point: 0x00009AF0;
Public Declare Function AboutDlgProc Lib "twain_32" () As Long

'Ordinal: 3(0x0003); Hint: 1(0x0001); Entry-Point: 0x00009600;
Public Declare Function ChooseDlgProc Lib "twain_32" () As Long

'Ordinal: 1(0x0001); Hint: 2(0x0002); Entry-Point: 0x0000B8F0;
Public Declare Function DSM_Entry Lib "twain_32" (pOriginSrc As TW_IDENTITY, pDestination As TW_IDENTITY, ByVal DG As Long, ByVal DAT As Integer, ByVal MSG As Integer, pData As TW_MEMORY) As Integer 'Long

Public Declare Function DS_Entry Lib "twain_32" (pOriginSrc As TW_IDENTITY, ByVal DG As Long, ByVal DAT As Integer, ByVal MSG As Integer, pData As TW_MEMORY) As Integer 'Long


'Ordinal: 6(0x0006); Hint: 3(0x0003); Entry-Point: 0x00009C60;
Public Declare Function InfoHook Lib "twain_32" () As Long

'Ordinal: 5(0x0005); Hint: 4(0x0004); Entry-Point: 0x00009C20;
Public Declare Function WGDlgProc Lib "twain_32" () As Long

'Ordinal: 2(0x0002); Hint: N/A;       Entry-Point: 0x00000000;

Public Function Single_ToTWFix32(ByVal Value As Single) As TW_FIX32
    Dim sign   As Boolean: sign = Value < 0
    Dim LngVal As Long:  LngVal = Value * 65536# + IIf(sign, -0.5, 0.5)
    With Single_ToTWFix32
        .Whole = Int(LngVal / 65536)
        .Frac = LngVal And &HFFFF&
    End With
End Function
'
'TW_FIX32 FloatToFIX32 (float floater)
'{
'  TW_FIX32 Fix32_value;
'  TW_BOOL  sign = (floater < 0)?TRUE:FALSE;
'  TW_INT32 value = (TW_INT32) (floater * 65536.0 + (sign?(-0.5):0.5));
'  Fix32_value.Whole = (TW_UINT16)(value >> 16);
'  Fix32_value.Frac = (TW_UINT16)(value & 0x0000ffffL);
'  return (Fix32_value);
'}

Public Function TWFix32_ToSingle(Value As TW_FIX32) As Single
    TWFix32_ToSingle = Value.Whole + Value.Frac / 65536#
End Function

Public Function getCurrent(pCap As TW_CAPABILITY, val_out As Long) As Boolean
    Dim bret As Boolean
    If pCap.hContainer = 0 Then Exit Function
End Function

Public Function TWAIN_Callback(pOrigin As TW_IDENTITY, ByVal DG As Long, ByVal DAT As Integer, ByVal MSG As Integer, pData As TW_MEMORY) As Integer
    '
End Function
