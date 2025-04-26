Imports System
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace TwainLib

	Public Class TwProtocol
		Public Const Major As Short = 1
		Public Const Minor As Short = 9
	End Class

#Region " Enums "
	<Flags()> Friend Enum TwDG As Short
		Control = &H1
		Image = &H2
		Audio = &H4
	End Enum

	Friend Enum TwDAT As Short
		Null = &H0
		Capability = &H1
		[Event] = &H2
		Identity = &H3
		Parent = &H4
		PendingXfers = &H5
		SetupMemXfer = &H6
		SetupFileXfer = &H7
		Status = &H8
		UserInterface = &H9
		XferGroup = &HA
		TwunkIdentity = &HB
		CustomDSData = &HC
		DeviceEvent = &HD
		FileSystem = &HE
		PassThru = &HF

		ImageInfo = &H101
		ImageLayout = &H102
		ImageMemXfer = &H103
		ImageNativeXfer = &H104
		ImageFileXfer = &H105
		CieColor = &H106
		GrayResponse = &H107
		RGBResponse = &H108
		JpegCompression = &H109
		Palette8 = &H10A
		ExtImageInfo = &H10B

		SetupFileXfer2 = &H301
	End Enum

	Friend Enum TwMSG As Short
		Null = &H0
		[Get] = &H1
		GetCurrent = &H2
		GetDefault = &H3
		GetFirst = &H4
		GetNext = &H5
		[Set] = &H6
		Reset = &H7
		QuerySupport = &H8

		XFerReady = &H101
		CloseDSReq = &H102
		CloseDSOK = &H103
		DeviceEvent = &H104

		CheckStatus = &H201

		OpenDSM = &H301
		CloseDSM = &H302

		OpenDS = &H401
		CloseDS = &H402
		UserSelect = &H403

		DisableDS = &H501
		EnableDS = &H502
		EnableDSUIOnly = &H503

		ProcessEvent = &H601

		EndXfer = &H701
		StopFeeder = &H702

		ChangeDirectory = &H801
		CreateDirectory = &H802
		Delete = &H803
		FormatMedia = &H804
		GetClose = &H805
		GetFirstFile = &H806
		GetInfo = &H807
		GetNextFile = &H808
		Rename = &H809
		Copy = &H80A
		AutoCaptureDir = &H80B

		PassThru = &H901
	End Enum

	Friend Enum TwRC As Short
		Success = &H0
		Failure = &H1
		CheckStatus = &H2
		Cancel = &H3
		DSEvent = &H4
		NotDSEvent = &H5
		XferDone = &H6
		EndOfList = &H7
		InfoNotSupported = &H8
		DataNotAvailable = &H9
	End Enum

	Friend Enum TwCC As Short
		Success = &H0
		Bummer = &H1
		LowMemory = &H2
		NoDS = &H3
		MaxConnections = &H4
		OperationError = &H5
		BadCap = &H6
		BadProtocol = &H9
		BadValue = &HA
		SeqError = &HB
		BadDest = &HC
		CapUnsupported = &HD
		CapBadOperation = &HE
		CapSeqError = &HF
		Denied = &H10
		FileExists = &H11
		FileNotFound = &H12
		NotEmpty = &H13
		PaperJam = &H14
		PaperDoubleFeed = &H15
		FileWriteError = &H16
		CheckDeviceOnline = &H17
	End Enum

	Friend Enum TwOn As Short
		Array = &H3
		[Enum] = &H4
		One = &H5
		Range = &H6
		DontCare = -1
	End Enum

	Friend Enum TwType As Short
		Int8 = &H0
		Int16 = &H1
		Int32 = &H2
		UInt8 = &H3
		UInt16 = &H4
		UInt32 = &H5
		Bool = &H6
		Fix32 = &H7
		Frame = &H8
		Str32 = &H9
		Str64 = &HA
		Str128 = &HB
		Str255 = &HC
		Str1024 = &HD
		Str512 = &HE
	End Enum

	Friend Enum TwCap As Short
		XferCount = &H1
		ICompression = &H100
		IPixelType = &H101
		IUnits = &H102
		IXferMech = &H103
	End Enum
#End Region

	<StructLayout(LayoutKind.Sequential, Pack:=2, CharSet:=CharSet.Ansi)> Friend Class TwIdentity
		Public Id As IntPtr		 'einzigartige Nummer, windows-handle
		Public Version As TwVersion
		Public ProtocolMajor As Short
		Public ProtocolMinor As Short
		Public SupportedGroups As Integer		 'unterstützte Daten (Image, Control, Audio, ...)
		<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=34)> Public Manufacturer As String
		<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=34)> Public ProductFamily As String
		<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=34)> Public ProductName As String
	End Class

	<StructLayout(LayoutKind.Sequential, Pack:=2, CharSet:=CharSet.Ansi)> Friend Structure TwVersion
		Public MajorNum As Short
		Public MinorNum As Short
		Public Language As Short
		Public Country As Short
		<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=34)> Public Info As String
	End Structure

	<StructLayout(LayoutKind.Sequential, Pack:=2)> Friend Class TwUserInterface
		Public ShowUI As Short	   '1 = UI zeigen, 0 = nicht zeigen
		Public ModalUI As Short	  'nur für MAC
		Public ParentHand As IntPtr	  'Fensterhandle
	End Class

	<StructLayout(LayoutKind.Sequential, Pack:=2)> Friend Class TwStatus
		Public ConditionCode As Short		 'Status, z.B. Success, Enum TwRC
		Public Reserved As Short	  'Space
	End Class

	<StructLayout(LayoutKind.Sequential, Pack:=2)> Friend Structure TwEvent
		Public EventPtr As IntPtr	   'Windows pMSG oder Mac pEvent
		Public Message As Short		 'message von Datenquelle, z.b. MSG_XFERREADY (Transferbereit)
	End Structure

	<StructLayout(LayoutKind.Sequential, Pack:=2)> Friend Class TwImageInfo
		Public XResolution As Integer
		Public YResolution As Integer
		Public ImageWidth As Integer
		Public ImageLength As Integer
		Public SamplesPerPixel As Short
		<MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)> Public BitsPerSample() As Short
		Public BitsPerPixel As Short
		Public Planar As Short
		Public PixelType As Short
		Public Compression As Short
	End Class

	<StructLayout(LayoutKind.Sequential, Pack:=2)> Friend Class TwPendingXfers	'prüft ob weitere Daten, zusammen mit MsgEndXfer
		Public Count As Short
		Public EOJ As Integer
	End Class

	<StructLayout(LayoutKind.Sequential, Pack:=2)> Friend Structure TwFix32
		Public Whole As Short
		Public Frac As Short

		Public Function ToFloat() As Single
			Return CType(Whole + (CType(Frac, Single) / 65536.0F), Single)
		End Function

		Public Sub FromFloat(ByVal f As Single)
			Dim i As Integer = CType(((f * 65536.0F) + 0.5F), Integer)
			Whole = CType(i / 2 ^ 16, Short)
			Frac = CType(i & &HFFFF, Short)
		End Sub
	End Structure

	<StructLayout(LayoutKind.Sequential, Pack:=2)> Friend Class TwCapability
		Public Cap As Short
		Public ConType As Short
		Public Handle As IntPtr
		Public Sub TwCapability(ByVal capIn As TwCap)
			Cap = CType(capIn, Short)
			ConType = -1
		End Sub

		Public Sub New(ByVal capIn As TwCap, ByVal sval As Short)
			Cap = CType(capIn, Short)
			ConType = CType(TwOn.One, Short)
			Handle = Twain.GlobalAlloc(&H42, 6)
			Dim pv As IntPtr = Twain.GlobalLock(Handle)
			Marshal.WriteInt16(pv, 0, CType(TwType.Int16, Short))
			Marshal.WriteInt32(pv, 2, CType(sval, Integer))
			Twain.GlobalUnlock(Handle)
		End Sub

		Public Sub Dispose()
			If Not Equals(Handle, IntPtr.Zero) Then
				Twain.GlobalFree(Handle)
			End If
		End Sub

		Protected Overrides Sub Finalize()
			If Not Equals(Handle, IntPtr.Zero) Then
				Twain.GlobalFree(Handle)
			End If
		End Sub
	End Class

End Namespace
