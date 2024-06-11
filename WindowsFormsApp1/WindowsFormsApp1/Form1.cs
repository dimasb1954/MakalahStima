using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataSet result;
        int len = 0;
        string query = null;
        int i = 0;
        
        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xls;*.xlsx", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    using (FileStream fs = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        IExcelDataReader reader = ExcelReaderFactory.CreateReader(fs);
                        result = reader.AsDataSet();
                        reader.Close();
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.query = textBox2.Text;
            bool success = int.TryParse(textBox1.Text, out this.len);

            if (!success)
            {
                MessageBox.Show("Please enter a valid integer for length.");
                return;
            }

            DataTable table = this.result.Tables[0];
            var rowsToDelete = new List<DataRow>();

            foreach (DataRow row in table.Rows)
            {
                bool rowMatches = false;

                foreach (var cell in row.ItemArray)
                {
                    string temp = cell.ToString();
                    if (temp.Contains(query) && temp.Length == len)
                    {
                        rowMatches = true;
                        break;
                    }
                }

                if (!rowMatches)
                {
                    rowsToDelete.Add(row);
                }
            }

            foreach (var row in rowsToDelete)
            {
                table.Rows.Remove(row);
            }

            dataGridView.DataSource = table;
        }
    }
}
