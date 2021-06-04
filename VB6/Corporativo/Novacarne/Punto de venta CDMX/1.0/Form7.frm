VERSION 5.00
Begin VB.Form Form7 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Vista Previa de Ticket ..."
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   2910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   2910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuImprimir 
      Caption         =   "Imprimir"
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuImprimir_Click()

    If PrImpresora <> "" Then
        Dim p As VB.Printer
        For Each p In VB.Printers
            If p.DeviceName = PrImpresora Then
                Set Printer = p
            End If
        Next
    End If

    Me.PrintForm

End Sub
