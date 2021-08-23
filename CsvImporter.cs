using System;

using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelBelegger
{
    class CsvImporter
    {
        private int rowIndex;
        private int columnIndex; // column A
        private Excel.Range SourceRange = null;

        public void openFile(String sheetName, System.IO.Stream fileStream)
        {
            var fileContent = string.Empty;

                using (StreamReader reader = new StreamReader(fileStream))
                {
                    string currentLine;
                    // currentLine will be null when the StreamReader reaches the end of file

                    Excel.Worksheet xlSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

                    xlSheet.Name = sheetName;


                    rowIndex = 1;


                    while ((currentLine = reader.ReadLine()) != null)
                    {

                        String[] seperated = Regex.Split(currentLine, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");//currentLine.Split(',');

                        //TODO make this generic
                        if (currentLine.Contains("geldmarktfonds"))
                            continue;



                        columnIndex = 65; // column A

                        foreach (String s in seperated)
                        {
                            char c = (char)columnIndex;

                            Excel.Range test = xlSheet.get_Range(c.ToString() + rowIndex);
                            test.Value2 = s;
                            columnIndex++;
                        }

                        rowIndex++;

                        fileContent += currentLine + "/n";


                    }

                    char check = (char)(columnIndex - 1);
                    SourceRange = (Excel.Range)xlSheet.get_Range("A1", check.ToString() + (rowIndex - 1)); // or whatever range you want here
                    FormatAsTable(SourceRange, sheetName, "TableStyleLight9");


                    //findAndHighlightValue(SourceRange, "koop", System.Drawing.Color.Green);
                    //findAndHighlightValue(SourceRange, "kosten", System.Drawing.Color.Red);
                    //findAndHighlightValue(SourceRange, "storting", System.Drawing.Color.Blue);
                    //findAndHighlightValue(SourceRange, "Valuta", System.Drawing.Color.Orange);
                    //findAndHighlightValue(SourceRange, "dividend", System.Drawing.Color.Purple);

                    //findAndReplaceValue(SourceRange, "\"", "");
                }
           
        }

        public Excel.Range GetRange()
        {
            return SourceRange;
        }

        private void FormatAsTable(Excel.Range SourceRange, string TableName, string TableStyleName)
        {
            // Check bouwen of tabel al bestaat
            if (SourceRange.Worksheet.ListObjects.Count > 0)
            {
                MessageBox.Show("Overschrijven bestaande tabel?", "Tabel bestaat al", MessageBoxButtons.YesNo);
                return;
            }
            else
            {
                SourceRange.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
                SourceRange, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing).Name =
                    TableName;
            }

            SourceRange.Select();
            SourceRange.Worksheet.ListObjects[TableName].TableStyle = TableStyleName;

        }

        public void findAndHighlightValue(Excel.Range searchRange, String searchTerm, System.Drawing.Color color)
        {
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;

            currentFind = searchRange.Find(searchTerm, Missing.Value,
            Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
            Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Missing.Value, Missing.Value);

            int i = 1;

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.Rows.Font.Color = System.Drawing.ColorTranslator.ToOle(color);
                currentFind.Rows.Font.Bold = true;

                currentFind = searchRange.FindNext(currentFind);

                i++;
            }
        }
    }
}
