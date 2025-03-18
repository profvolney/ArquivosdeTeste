using Microsoft.Office.Interop.Excel;

string filePath = @"C:\\Users\\jason.vsilva\\Desktop\\ReadExcel\\planilha.xls";


var excel = new Application();

Workbook wb;
Worksheet ws;

wb = excel.Workbooks.Open(filePath);
ws = (Worksheet)wb.Worksheets[1];

Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)ws.Cells[3,2];
Microsoft.Office.Interop.Excel.Range cell1 = (Microsoft.Office.Interop.Excel.Range)ws.Cells[3,2];

string CellValue = (string)cell.Value;
string Cell1Value = (string)cell1.Value;

Console.WriteLine("O conteúdo da coluna é " 
+ CellValue + " e a célula é " + Cell1Value);
