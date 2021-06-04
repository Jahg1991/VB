Attribute VB_Name = "Module1"
Global CnDb As New ADODB.Connection

Global RsUsers As New ADODB.Recordset
Global RsSales As New ADODB.Recordset
Global RSReports As New ADODB.Recordset
Global RSTotals As New ADODB.Recordset
Global RSYearWeek As New ADODB.Recordset

Global STName As String
Global STProm As String


Sub Main()
    
    With CnDb
        .CursorLocation = adUseClient
        .Open "Provider=SQLOLEDB.1;Password=Jahg1991;Persist Security Info=True;User ID=sa;Initial Catalog=PIGSALE;Data Source=localhost\SSDB;"
    End With
    
    Form1.Show

End Sub
