using Microsoft.EntityFrameworkCore.Query;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestQ
{
    public partial class Form1 : Form
    {
        private Excel.Application ExcelApp = new Excel.Application();
        public string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database1.mdb";
        private OleDbConnection connect_db;
        private OleDbCommand comand;
        private DataTable dt;
        private OleDbDataAdapter adapt;
        string filepath;
        Excel.Workbook ExlWorkBook;
        Excel.Worksheet ExlWorkSheet;
        public Form1()
        {
            InitializeComponent();
            OleDbDataAdapter adapter = new OleDbDataAdapter();
            connect_db = new OleDbConnection(connectionString);
            connect_db.Open();
            comand = new OleDbCommand();
            adapt = new OleDbDataAdapter("SELECT * FROM process_table", connect_db);
            OleDbCommandBuilder cBuilder = new OleDbCommandBuilder(adapt);
            dt = new DataTable();

            adapt.Fill(dt);

            dataGridView1.DataSource = dt;

            //dataGridView1.Visible = false;
            //dataGridView2.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int RowsCount;
            int ColumnsCount;
            string[] process_id;
            string[] name_process;
            string[] dep_name;


            try
            {
                ExlWorkBook = ExcelApp.Workbooks.Open(filepath);
                ExlWorkSheet = ExlWorkBook.Worksheets[1];
                var arrayData = (object[,])ExlWorkSheet.Range["B4:E200"].Value;
                ExlWorkBook.Close();
                ExcelApp.Quit();
                this.dataGridView2.Rows.Clear();
                RowsCount = arrayData.GetUpperBound(0);
                ColumnsCount = arrayData.GetUpperBound(1);
                process_id = new string[RowsCount];
                name_process = new string[RowsCount];
                dep_name = new string[RowsCount];
                dataGridView2.RowCount = RowsCount;
                dataGridView2.ColumnCount = ColumnsCount;
                int i, j;
                for (i = 1; i < RowsCount; i++)
                {
                    for (j = 1; j < ColumnsCount; j++)
                    {
                        if (arrayData[i, j] != null)
                        {
                            dataGridView2.Rows[i - 1].Cells[j - 1].Value = arrayData[i, j];
                        }

                    }
                }
                for (i = 1; i < dataGridView2.Rows.Count; i++)
                {
                    if (dataGridView2.Rows[i].Cells[0].Value != null) 
                    {
                        process_id[i] = dataGridView2.Rows[i].Cells[0].Value.ToString();
                    }
                    if (dataGridView2.Rows[i].Cells[1].Value != null) 
                    {
                        name_process[i] = dataGridView2.Rows[i].Cells[1].Value.ToString();
                    }
                    if (dataGridView2.Rows[i].Cells[2].Value != null) 
                    {
                        dep_name[i] = dataGridView2.Rows[i].Cells[2].Value.ToString();
                    }
                }

                for (i = 1; i<RowsCount; i++)
                {
                    dt.Rows.Add(process_id[i], name_process[i], dep_name[i]);
                }
                adapt.Update(dt);
                label1.Text = "Данные записаны в базу!";
                connect_db.Close();
            }
            catch(Exception ex) {
                label1.Text = "Ошибка открытия файла";
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK) 
            {
                filepath = ofd.FileName;
            }
        }
    }
}
