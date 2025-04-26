Imports System
Imports System.Collections
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Namespace TwainLib

	Public Enum TwainCommand
		[Not] = -1
		Null = 0
		TransferReady = 1
		CloseRequest = 2
		CloseOk = 3
		DeviceEvent = 4
	End Enum

	Public Class Twain

		Private hwnd As IntPtr
		Private appid As TwIdentity
		Private srcds As TwIdentity
		Private evtmsg As TwEvent
		Private winmsg_m As WINMSG_S

#Region " API Import "
		<DllImport("twain_32.dll", EntryPoint:="#1")> Private Shared Function DSMparent(<[In](), Out()> ByVal origin As TwIdentity, ByVal zeroptr As IntPtr, ByVal dg As TwDG, ByVal dat As TwDAT, ByVal msg As TwMSG, ByRef refptr As IntPtr) As TwRC
		End Function

		<DllImport("twain_32.dll", EntryPoint:="#1")> Private Shared Function DSMident(<[In](), Out()> ByVal origin As TwIdentity, ByVal zeroptr As IntPtr, ByVal dg As TwDG, ByVal dat As TwDAT, ByVal msg As TwMSG, <[In](), Out()> ByVal idds As TwIdentity) As TwRC
		End Function

		<DllImport("twain_32.dll", EntryPoint:="#1")> Private Shared Function DSMstatus(<[In](), Out()> ByVal origin As TwIdentity, ByVal zeroptr As IntPtr, ByVal dg As TwDG, ByVal dat As TwDAT, ByVal msg As TwMSG, <[In](), Out()> ByVal dsmstat As TwStatus) As TwRC
		End Function

		<DllImport("twain_32.dll", EntryPoint:="#1")> Private Shared Function DSuserif(<[In](), Out()> ByVal origin As TwIdentity, <[In](), Out()> ByVal dest As TwIdentity, ByVal dg As TwDG, ByVal dat As TwDAT, ByVal msg As TwMSG, ByVal guif As TwUserInterface) As TwRC
		End Function

		<DllImport("twain_32.dll", EntryPoint:="#1")> Private Shared Function DSevent(<[In](), Out()> ByVal origin As TwIdentity, <[In](), Out()> ByVal dest As TwIdentity, ByVal dg As TwDG, ByVal dat As TwDAT, ByVal msg As TwMSG, ByRef evt As TwEvent) As TwRC
		End Function

		<DllImport("twain_32.dll", EntryPoint:="#1")> Private Shared Function DSstatus(<[In](), Out()> ByVal origin As TwIdentity, <[In]()> ByVal dest As TwIdentity, ByVal dg As TwDG, ByVal dat As TwDAT, ByVal msg As TwMSG, <[In](), Out()> ByVal dsmstat As TwStatus) As TwRC
		End Function

		<DllImport("twain_32.dll", EntryPoint:="#1")> Private Shared Function DScap(<[In](), Out()> ByVal origin As TwIdentity, <[In]()> ByVal dest As TwIdentity, ByVal dg As TwDG, ByVal dat As TwDAT, ByVal msg As TwMSG, <[In](), Out()> ByVal capa As TwCapability) As TwRC
		End Function

		<DllImport("twain_32.dll", EntryPoint:="#1")> Private Shared Function DSiinf(<[In](), Out()> ByVal origin As TwIdentity, <[In]()> ByVal dest As TwIdentity, ByVal dg As TwDG, ByVal dat As TwDAT, ByVal msg As TwMSG, <[In](), Out()> ByVal imginf As TwImageInfo) As TwRC
		End Function

		<DllImport("twain_32.dll", EntryPoint:="#1")> Private Shared Function DSixfer(<[In](), Out()> ByVal origin As TwIdentity, <[In]()> ByVal dest As TwIdentity, ByVal dg As TwDG, ByVal dat As TwDAT, ByVal msg As TwMSG, ByRef hbitmap As IntPtr) As TwRC
		End Function

		<DllImport("twain_32.dll", EntryPoint:="#1")> Private Shared Function DSpxfer(<[In](), Out()> ByVal origin As TwIdentity, <[In]()> ByVal dest As TwIdentity, ByVal dg As TwDG, ByVal dat As TwDAT, ByVal msg As TwMSG, <[In](), Out()> ByVal pxfr As TwPendingXfers) As TwRC
		End Function

		<DllImport("kernel32.dll", ExactSpelling:=True)> Friend Shared Function GlobalAlloc(ByVal flags As Integer, ByVal size As Integer) As IntPtr
		End Function

		<DllImport("kernel32.dll", ExactSpelling:=True)> Friend Shared Function GlobalLock(ByVal handle As IntPtr) As IntPtr
		End Function

		<DllImport("kernel32.dll", ExactSpelling:=True)> Friend Shared Function GlobalUnlock(ByVal handle As IntPtr) As Boolean
		End Function

		<DllImport("kernel32.dll", ExactSpelling:=True)> Friend Shared Function GlobalFree(ByVal handle As IntPtr) As IntPtr
		End Function

		<DllImport("user32.dll", ExactSpelling:=True)> Private Shared Function GetMessagePos() As Integer
		End Function

		<DllImport("user32.dll", ExactSpelling:=True)> Private Shared Function GetMessageTime() As Integer
		End Function

		<DllImport("gdi32.dll", ExactSpelling:=True)> Private Shared Function GetDeviceCaps(ByVal hDC As IntPtr, ByVal nIndex As Integer) As Integer
		End Function

		<DllImport("gdi32.dll", CharSet:=CharSet.Auto)> Private Shared Function CreateDC(ByVal szdriver As String, ByVal szdevice As String, ByVal szoutput As String, ByVal devmode As IntPtr) As IntPtr
		End Function

		<DllImport("gdi32.dll", ExactSpelling:=True)> Private Shared Function DeleteDC(ByVal hdc As IntPtr) As Boolean
		End Function
#End Region

		Private Const CountryUSA As Short = 1
		Private Const LanguageUSA As Short = 13

		Public Sub New()
			appid = New TwIdentity
			appid.Id = IntPtr.Zero
			appid.Version.MajorNum = 1
			appid.Version.MinorNum = 1
			appid.Version.Language = LanguageUSA
			appid.Version.Country = CountryUSA
			appid.Version.Info = "Converted"
			appid.ProtocolMajor = TwProtocol.Major
			appid.ProtocolMinor = TwProtocol.Minor
			appid.SupportedGroups = CType(TwDG.Image Or TwDG.Control, Integer)
			appid.Manufacturer = "MF@mcs"
			appid.ProductFamily = "Freeware"
			appid.ProductName = "Converted Twain"

			srcds = New TwIdentity
			srcds.Id = IntPtr.Zero

			evtmsg.EventPtr = Marshal.AllocHGlobal(Marshal.SizeOf(winmsg_m))
		End Sub

		Public Sub Dispose()
			Marshal.FreeHGlobal(evtmsg.EventPtr)
		End Sub

		Protected Overrides Sub Finalize()
			Marshal.FreeHGlobal(evtmsg.EventPtr)
		End Sub

		Public Sub Init(ByVal hwndp As IntPtr)
			Finish()
			Dim rc As TwRC = DSMparent(appid, IntPtr.Zero, TwDG.Control, TwDAT.Parent, TwMSG.OpenDSM, hwndp)
			If (rc = TwRC.Success) Then
				rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.GetDefault, srcds)
				If (rc = TwRC.Success) Then
					hwnd = hwndp
				Else
					rc = DSMparent(appid, IntPtr.Zero, TwDG.Control, TwDAT.Parent, TwMSG.CloseDSM, hwndp)
				End If
			End If
		End Sub

		Public Sub [Select]()
			Dim rc As TwRC
			CloseSrc()
			If Equals(appid.Id, IntPtr.Zero) = True Then
				Init(hwnd)
				If Equals(appid.Id, IntPtr.Zero) = True Then
					Return
				End If
			End If
			rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.UserSelect, srcds)
		End Sub

		Public Sub Acquire()
			Dim rc As TwRC
			CloseSrc()
			If Equals(appid.Id, IntPtr.Zero) = True Then
				Init(hwnd)
				If Equals(appid.Id, IntPtr.Zero) = True Then
					Return
				End If
			End If

			rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.OpenDS, srcds)
			If (rc <> TwRC.Success) Then
				Return
			End If

			Dim cap As TwCapability = New TwCapability(TwCap.XferCount, 1)
			rc = DScap(appid, srcds, TwDG.Control, TwDAT.Capability, TwMSG.Set, cap)
			If (rc <> TwRC.Success) Then
				CloseSrc()
				Return
			End If

			Dim guif As TwUserInterface = New TwUserInterface
			guif.ShowUI = 1		   'Optionen zum scannen anzeigen, 0 = sofort scannen
			guif.ModalUI = 1
			guif.ParentHand = hwnd
			rc = DSuserif(appid, srcds, TwDG.Control, TwDAT.UserInterface, TwMSG.EnableDS, guif)
			If (rc <> TwRC.Success) Then
				CloseSrc()
				Return
			End If
		End Sub

		Public Function TransferPictures() As ArrayList
			Dim pics As ArrayList = New ArrayList
			If Equals(srcds.Id, IntPtr.Zero) Then
				Return pics
			End If

			Dim rc As TwRC
			Dim hbitmap As IntPtr = IntPtr.Zero
			Dim pxfr As TwPendingXfers = New TwPendingXfers

			Do
				pxfr.Count = 0
				hbitmap = IntPtr.Zero

				Dim iinf As TwImageInfo = New TwImageInfo
				rc = DSiinf(appid, srcds, TwDG.Image, TwDAT.ImageInfo, TwMSG.Get, iinf)
				If (rc <> TwRC.Success) Then
					CloseSrc()
					Return pics
				End If

				rc = DSixfer(appid, srcds, TwDG.Image, TwDAT.ImageNativeXfer, TwMSG.Get, hbitmap)
				If (rc <> TwRC.XferDone) Then
					CloseSrc()
					Return pics
				End If

				rc = DSpxfer(appid, srcds, TwDG.Control, TwDAT.PendingXfers, TwMSG.EndXfer, pxfr)
				If (rc <> TwRC.Success) Then
					CloseSrc()
					Return pics
				End If

				pics.Add(hbitmap)
			Loop While (pxfr.Count <> 0)

			rc = DSpxfer(appid, srcds, TwDG.Control, TwDAT.PendingXfers, TwMSG.Reset, pxfr)
			Return pics
		End Function

		Public Function PassMessage(ByVal m As Message) As TwainCommand
			If Equals(srcds.Id, IntPtr.Zero) Then
				Return TwainCommand.Not
			End If

			Dim pos As Integer = GetMessagePos()

			winmsg_m.hwnd = m.HWnd
			winmsg_m.message = m.Msg
			winmsg_m.wParam = m.WParam
			winmsg_m.lParam = m.LParam
			winmsg_m.time = GetMessageTime()
			winmsg_m.x = pos			 'CType(pos, Short)
			winmsg_m.y = Int(pos / 2 ^ 16)			 'CType(Int(pos / 2 ^ 16), Short)

			Marshal.StructureToPtr(winmsg_m, evtmsg.EventPtr, False)
			evtmsg.Message = 0

			Dim rc As TwRC = DSevent(appid, srcds, TwDG.Control, TwDAT.Event, TwMSG.ProcessEvent, evtmsg)
			If (rc = TwRC.NotDSEvent) Then
				Return TwainCommand.Not
			End If

			If (evtmsg.Message = CType(TwMSG.XFerReady, Short)) Then
				Return TwainCommand.TransferReady
			End If

			If (evtmsg.Message = CType(TwMSG.CloseDSReq, Short)) Then
				Return TwainCommand.CloseRequest
			End If
			If (evtmsg.Message = CType(TwMSG.CloseDSOK, Short)) Then
				Return TwainCommand.CloseOk
			End If
			If (evtmsg.Message = CType(TwMSG.DeviceEvent, Short)) Then
				Return TwainCommand.DeviceEvent
			End If

			Return TwainCommand.Null
		End Function

		Public Sub CloseSrc()
			Dim rc As TwRC
			If Not Equals(srcds.Id, IntPtr.Zero) Then
				Dim guif As TwUserInterface = New TwUserInterface
				rc = DSuserif(appid, srcds, TwDG.Control, TwDAT.UserInterface, TwMSG.DisableDS, guif)
				rc = DSMident(appid, IntPtr.Zero, TwDG.Control, TwDAT.Identity, TwMSG.CloseDS, srcds)
			End If
		End Sub

		Public Sub Finish()
			Dim rc As TwRC
			CloseSrc()
			If Not Equals(appid.Id, IntPtr.Zero) Then
				rc = DSMparent(appid, IntPtr.Zero, TwDG.Control, TwDAT.Parent, TwMSG.CloseDSM, hwnd)
			End If
			appid.Id = IntPtr.Zero
		End Sub

		Public Shared ReadOnly Property ScreenBitDepth() As Integer
			Get
				Dim screenDC As IntPtr = CreateDC("DISPLAY", Nothing, Nothing, IntPtr.Zero)
				Dim bitDepth As Integer = GetDeviceCaps(screenDC, 12)
				bitDepth *= GetDeviceCaps(screenDC, 14)
				DeleteDC(screenDC)
				Return bitDepth
			End Get
		End Property

		<StructLayout(LayoutKind.Sequential, Pack:=4)> Friend Structure WINMSG_S
			Public hwnd As IntPtr
			Public message As Integer
			Public wParam As IntPtr
			Public lParam As IntPtr
			Public time As Integer
			Public x As Integer
			Public y As Integer
		End Structure
	End Class
End Namespace
