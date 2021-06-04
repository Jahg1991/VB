Attribute VB_Name = "Module2"
Option Explicit


'Constantes
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const PD_DISABLEPRINTTOFILE = &H80000

'Funciones API
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" ( _
    pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" ( _
    pPagesetupdlg As PAGESETUPDLG) As Long
Private Declare Function GlobalLock Lib "kernel32" ( _
    ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" ( _
    ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    hpvDest As Any, _
    hpvSource As Any, _
    ByVal cbCopy As Long)

' UDT
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PAGESETUPDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    Flags As Long
    ptPaperSize As POINTAPI
    rtMinMargin As RECT
    rtMargin As RECT
    hInstance As Long
    lCustData As Long
    lpfnPageSetupHook As Long
    lpfnPagePaintHook As Long
    lpPageSetupTemplateName As String
    hPageSetupTemplate As Long
End Type


Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'Fin de declaraciones

'----------------------------------


' función Para el Common diálogo de Configurar página
'---------------------------------------------------------
Function Configuarar_Pagina(HwndForm As Long) As Long

    Dim T_Configurar_Pagina As PAGESETUPDLG
    
    
    With T_Configurar_Pagina
        .lStructSize = Len(T_Configurar_Pagina)
        .hwndOwner = HwndForm
        .hInstance = App.hInstance
        .Flags = 0
    End With
    
    If PAGESETUPDLG(T_Configurar_Pagina) Then
        Configuarar_Pagina = 0
    Else
        Configuarar_Pagina = -1
    End If
End Function


'Para el Common diálogo de imprimir ( pasar el formulario como parámetro )
'---------------------------------------------------------
Public Sub Show_Printer(El_Formulario As Form, Optional Flags As Long)
    
    On Error GoTo ErrSub
    
    Dim t_Printer As PRINTDLG_TYPE
    Dim DevMode As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    With t_Printer
        .lStructSize = Len(t_Printer)
        .hwndOwner = El_Formulario.hWnd
        .Flags = Flags
    End With
    
    On Error Resume Next
    
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex
    
    On Error GoTo 0
    
    t_Printer.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(t_Printer.hDevMode)
    
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
        bReturn = GlobalUnlock(t_Printer.hDevMode)
    End If

    
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    t_Printer.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(t_Printer.hDevNames)
    
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    
    If PrintDialog(t_Printer) <> 0 Then
        lpDevName = GlobalLock(t_Printer.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree t_Printer.hDevNames

        
        lpDevMode = GlobalLock(t_Printer.hDevMode)
        CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
        bReturn = GlobalUnlock(t_Printer.hDevMode)
        GlobalFree t_Printer.hDevMode
        NewPrinterName = UCase$(Left(DevMode.dmDeviceName, _
                InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
        
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                End If
            Next
        End If

        On Error Resume Next
        
        With Printer
            .PaperSize = DevMode.dmPaperSize
            .PrintQuality = DevMode.dmPrintQuality
            .ColorMode = DevMode.dmColor
            .PaperBin = DevMode.dmDefaultSource
            .Copies = DevMode.dmCopies
            .Duplex = DevMode.dmDuplex
            .Orientation = DevMode.dmOrientation
        End With
        On Error GoTo 0
    
    End If

Exit Sub

ErrSub:

If Err.Number = 484 Then
    MsgBox "Error al obtener información de la impresora." & _
            "Asegurarse que está instalada correctamente.", vbCritical
End If

End Sub


