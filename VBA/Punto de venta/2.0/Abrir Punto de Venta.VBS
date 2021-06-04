Public AplicacionExcel
Public AplicacionExcelTemp
Set AplicacionExcelTemp = CreateObject("Excel.Application")
AplicacionExcelTemp.Visible = True
AplicacionExcelTemp.Workbooks.Add()
Set AplicacionExcel = CreateObject("Excel.Application")
AplicacionExcel.Visible = True
AplicacionExcel.Workbooks.Open ("\\SRVDOM\Facturacion\Punto de venta v2\Punto de venta\Punto de Venta.xlsm")
