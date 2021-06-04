Attribute VB_Name = "Ingreso"
Option Explicit
  
Public Sub Main()
On Error Resume Next
  
  
    ' Abre el formulario para el ingreso _
      del Usuario y la contraseña
    frmLogin.Show vbModal
  
    MsgBox "Bienvenido(a)  ", vbInformation, " Login Correcto "
    ' Abre el formulario principal del programa
    frmSplash.Show
  
End Sub

