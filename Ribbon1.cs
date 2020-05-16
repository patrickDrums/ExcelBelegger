using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelBelegger
{
    public partial class Ribbon1
    {
        private int rowIndex;
        private int columnIndex; // column A


        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        public void FormatAsTable(Excel.Range SourceRange, string TableName, string TableStyleName)
        {
            // Check bouwen of tabel al bestaat
            
            SourceRange.Worksheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
            SourceRange, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing).Name =
                TableName;
            SourceRange.Select();
            SourceRange.Worksheet.ListObjects[TableName].TableStyle = TableStyleName;
        }

        private void createDividendPivotTable(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet xlSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            Excel.Range currentFind = null;
            Excel.Range firstFind = null;

            if(!xlSheet.Name.Equals("Account"))
            {
                MessageBox.Show("Selecteer het 'Account' tabblad.");
                return;
            }
            char c = 'P';

            Excel.Range accountTable = xlSheet.ListObjects["Account"].Range;

            Excel.Range dividendTable;


            currentFind = accountTable.Find("dividend", Missing.Value,
           Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
           Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false, Missing.Value, Missing.Value);

            int i =1;

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

               dividendTable  = xlSheet.get_Range(c.ToString() + i);


                MessageBox.Show("Cell: " + currentFind.Column + " : " + currentFind.Row);
               
                
                currentFind.Rows.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                currentFind.Rows.Font.Bold = true;

                dividendTable.Value2 = currentFind.Value2;

                currentFind = accountTable.FindNext(currentFind);

                i++;
            }


            
        }


        private void loadAccountData(object sender, RibbonControlEventArgs e)
        {
            var fileContent = string.Empty;
            var filePath = string.Empty;

            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the path of specified file
                filePath = openFileDialog.FileName;

                //Read the contents of the file into a stream
                var fileStream = openFileDialog.OpenFile();

                using (StreamReader reader = new StreamReader(fileStream))
                {
                    string currentLine;
                    // currentLine will be null when the StreamReader reaches the end of file

                    



                    Excel.Worksheet xlSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                    xlSheet.Name = "Account";


                    rowIndex = 1;

                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        


                        String[] seperated = Regex.Split(currentLine, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");//currentLine.Split(',');


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

                    MessageBox.Show(fileContent, "File Content at path: " + filePath, MessageBoxButtons.OK);

                    char check = (char) (columnIndex-1);
                    Excel.Range SourceRange = (Excel.Range)xlSheet.get_Range("A1", check.ToString() + (rowIndex-1)); // or whatever range you want here
                    FormatAsTable(SourceRange, "Account", "TableStyleLight9");
                }
            }
        }
    }
}
