using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ClosedXML;
using ClosedXML.Excel;
using System.IO;

namespace QueryGen
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProcessStart();
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (openFileDialog1.FileName.ToLower().EndsWith(".xls") || openFileDialog1.FileName.ToLower().EndsWith(".xlsx"))
                {
                    comboBox1.Items.Clear();

                    var wb = new XLWorkbook(openFileDialog1.FileName);

                    List<IXLWorksheet> sheets = wb.Worksheets.ToList<IXLWorksheet>();

                    foreach (IXLWorksheet s in sheets)
                    {
                        comboBox1.Items.Add(s.Name);
                    }



                }
                else
                {
                    MessageBox.Show("Select a excel file !");
                }
            }
            ProcessEnd();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ProcessStart();

            if (ValidateUI())
            {
                QueryBuilderInput qbInput = CreateRequest();

                string query = QueryProcessor.GenerateQuery(qbInput);

                txtQuery.Text = query;
            }
            ProcessEnd();
        }


        private bool ValidateUI()
        {
            bool isValid = true;

            if (comboBox1.SelectedItem == null || comboBox1.SelectedItem == string.Empty)
            {
                isValid = false;
                MessageBox.Show("Please select a excel sheet.");

            }

            if (string.IsNullOrEmpty(txtTableName.Text))
            {
                isValid = false;
                MessageBox.Show("Please enter a DB table name.");
            }

            try
            {
                Convert.ToInt32(txtFrom.Text);
            }
            catch (Exception ex)
            {
                isValid = false;
                MessageBox.Show("Please enter a valid number in From textbox.");
            }

            try
            {
                Convert.ToInt32(txtTo.Text);
            }
            catch (Exception ex)
            {
                isValid = false;
                MessageBox.Show("Please enter a valid number in To textbox.");
            }

            return isValid;
        }


        private QueryBuilderInput CreateRequest()
        {
            List<Column> collCollection = new List<Column>();

            Column col;
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                if (dr.Cells["ExcelColumn"].Value != null && dr.Cells["TableColumn"].Value != null)
                {
                    col = new Column();
                    col.ColumnNameExcel = dr.Cells["ExcelColumn"].Value.ToString();
                    col.ColumnNameDB = dr.Cells["TableColumn"].Value.ToString();
                    col.IsString = !(Convert.ToBoolean(dr.Cells["IsNumeric"].Value));
                    col.PassNull = (Convert.ToBoolean(dr.Cells["PassNull"].Value));
                    collCollection.Add(col);
                }
            }

            QueryBuilderInput qbInput = new QueryBuilderInput();
            qbInput.SheetName = comboBox1.SelectedItem.ToString();
            qbInput.StartRow = Convert.ToInt32(txtFrom.Text);
            qbInput.EndRow = Convert.ToInt32(txtTo.Text);
            qbInput.ExcelColumns = collCollection;
            qbInput.TableNameDB = txtTableName.Text;
            qbInput.File = openFileDialog1.FileName;
            return qbInput;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ProcessStart();
            if (!string.IsNullOrEmpty(txtQuery.Text))
            {
                string validationMessage = "";
                bool hasErrors = QueryProcessor.ValidateQuery(txtQuery.Text.Replace("\r\n", ""), out validationMessage);

                if (!hasErrors)
                    MessageBox.Show("Query is valid...");
                else
                {
                    MessageBox.Show(validationMessage);
                    MessageBox.Show("Validation fails even when query looks fine? Better add [ ] for column names to avoid parsing errors in case of using reserved keywords for column names.");

                }
            }
            else
            {
                MessageBox.Show("Query is empty. Create scripts first...");
            }
            ProcessEnd();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ProcessStart();
            if (!string.IsNullOrEmpty(txtQuery.Text))
            {

                if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    File.WriteAllText(saveFileDialog1.FileName, txtQuery.Text);
                    MessageBox.Show("Query saved at " + saveFileDialog1.FileName);
                }
            }
            else
            {
                MessageBox.Show("Query is empty. Create scripts first...");
            }
            ProcessEnd();
        }


        private void ProcessStart()
        {
            toolStripStatusLabel1.Text = "Processing ...";
            statusStrip1.Refresh();
        }

        private void ProcessEnd()
        {
            toolStripStatusLabel1.Text = "Ready";
            statusStrip1.Refresh();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ProcessStart();
            if (!string.IsNullOrEmpty(txtQuery.Text))
            {
                System.Windows.Forms.Clipboard.SetText(txtQuery.Text);
            }
            else
            {
                MessageBox.Show("Query is empty. Create scripts first...");
            }
            ProcessEnd();


        }
    }
}
