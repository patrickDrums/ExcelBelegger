using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelBelegger
{
    public partial class FileChooser : Form
    {
        public FileChooser()
        {
            InitializeComponent();
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void browseBtn_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ShowDialog();
        }

        private void FileChooser_Load(object sender, EventArgs e)
        {
            List<KeyValuePair<string, string>> transactionList = new List<KeyValuePair<string, string>>();
            Array transactions = Enum.GetValues(typeof(CsvTransactions));
            foreach (CsvTransactions transaction in transactions)
            {
                transactionList.Add(new KeyValuePair<string, string>(transaction.ToString(), ((int)transaction).ToString()));
            }
            comboBox1.DataSource = transactionList;
            comboBox1.DisplayMember = "Key";
            comboBox1.ValueMember = "Value";
        }

        private void OkBtn_Click(object sender, EventArgs e)
        {
            CsvImporter csvImporter = new CsvImporter();

            //Read the contents of the file into a stream
            csvImporter.openFile(Enum.GetName(typeof(CsvTransactions), comboBox1.SelectedIndex), openFileDialog1.OpenFile());

            this.Close();
        }
    }
}
