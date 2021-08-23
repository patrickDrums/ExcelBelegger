using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        } 

        public void findAndReplaceValue(Excel.Range searchRange, String searchTerm, String replacement)
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

                currentFind.Value2 = Regex.Replace(currentFind.Value2, searchTerm, replacement);

                
                currentFind = searchRange.FindNext(currentFind);

                i++;
            }
        }

        private void createDividendPivotTable(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet xlSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            String[] columnNames;

            if(!xlSheet.Name.Equals("Account"))
            {
                MessageBox.Show("Selecteer eerst het 'Account' tabblad.");
                return;
            }

            Excel.Range accountTable = xlSheet.ListObjects["Account"].Range;
            Excel.Range headerRows = accountTable.Rows[1]; // first row

            columnNames = new String[12];

            int i = 0;

            foreach (Excel.Range item in headerRows.Cells)
            {
                columnNames[i] = String.Empty;
                columnNames[i] = item.Value2;

                i++;
            }

            SelectColumnForm form = new SelectColumnForm();
            form.setCollectionComboBoxes(columnNames.ToArray());
            form.ShowDialog();

            int dateIndex = 0;
            int productIndex = 0;
            int descriptionIndex = 0;
            int saldoIndex = 0;

            if (form.DialogResult == DialogResult.OK)
            {
                dateIndex = (char) form.getSelectedIndexDateColumn();
                productIndex = (char) form.getSelectedIndexProductColumn();
                descriptionIndex = (char)form.getSelectedIndexDescriptionColumn();
                saldoIndex = (char)form.getSelectedIndexSaldoColumn();


                // Optional: Call the Dispose method when you are finished with the dialog box.
                form.Dispose();
            }

            MessageBox.Show(columnNames[dateIndex] + ", " + columnNames[productIndex] + ", " + columnNames[descriptionIndex] + ", " + columnNames[saldoIndex]);

            //Copy to new Table
            //accountTable.Columns[dateIndex + 1].Copy();
            //xlSheet.get_Range("P1").Select();
            //xlSheet.Paste();
            //accountTable.Columns[productIndex + 1].Copy();
            //xlSheet.get_Range("Q1").Select();
            //xlSheet.Paste();
            //accountTable.Columns[descriptionIndex + 1].Copy();
            //xlSheet.get_Range("R1").Select();
            //xlSheet.Paste();
            //accountTable.Columns[saldoIndex + 1].Copy();
            //xlSheet.get_Range("S1").Select();
            //xlSheet.Paste();


            //accountTable.Columns[productIndex + 1].Select();
            //accountTable.Columns[descriptionIndex + 1].Select();
            //accountTable.Columns[saldoIndex + 1].Select();

            

        }

        private void loadAccountData(object sender, RibbonControlEventArgs e)
        {
            String[] highlights = new string[] { "koop", "kosten", "storting", "Valuta", "dividend" };
            CsvImporter importer = new CsvImporter();

            FileChooser fc = new FileChooser();
            fc.Show();

            //importer.openFile("Account");

            //Excel.Range tableRange = importer.GetRange();


            //importer.findAndHighlightValue(tableRange, highlights[0], System.Drawing.Color.Green);

            //findAndHighlightValue(SourceRange, "koop", System.Drawing.Color.Green);
            //findAndHighlightValue(SourceRange, "kosten", System.Drawing.Color.Red);
            //findAndHighlightValue(SourceRange, "storting", System.Drawing.Color.Blue);
            //findAndHighlightValue(SourceRange, "Valuta", System.Drawing.Color.Orange);
            //findAndHighlightValue(SourceRange, "dividend", System.Drawing.Color.Purple);

            //findAndReplaceValue(SourceRange, "\"", "");
            
        }

        private void loadCryptoFiat(object sender, RibbonControlEventArgs e)
        {
            String[] highlights = new string[] { "->", "Deposit", "Withdrawal", "Top" };
            
        }

        private void loadCrypto(object sender, RibbonControlEventArgs e)
        {
            String[]highlights = new string[] {"Card", "Rebate"};
            
        }

        private void loadCoinmerce(object sender, RibbonControlEventArgs e)
        {
            String[] highlights = new string[] { };

        }
    }
}
