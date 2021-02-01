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

namespace TestCheckInOffice
{
    public partial class Form1 : Form
    {
        DataSet result;
        DataRowCollection dataRow;
        DataColumnCollection dataColumn;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                DialogResult ans = MessageBox.Show("確定匯入資料?", "訊息視窗",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (ans == DialogResult.OK)
                {
                    ImportFile();
                }
                else
                {
                    return;
                }

            }
            catch (Exception ex)
            {
                string error = ex.ToString();
                MessageBox.Show(error, "執行錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        protected void ImportFile()
        {
            OpenFileDialog dialog = new OpenFileDialog(); //建立檔案選擇視窗
            dialog.Title = "Please your files";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "xls Files(*.xls; *.xlsx;)| *.xls; *.xlsx;";
            //"xls Files(*.xls; *.xlsx;)| *.xls; *.xlsx; | All files(*.*) | *.*"
            string msg = "";
            string xlsfile = "";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                xlsfile = dialog.FileName;
                msg = "Are you sure import " + dialog.FileName + " ?";
                DialogResult ans = MessageBox.Show(msg, "Check Message",
                                   MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (ans == DialogResult.OK)
                {

                    //2.抓取Excel內容
                    DataSet dtXls = ExcelProcess(xlsfile);
       
                    MessageBox.Show("Import Compelete!");
                }
            }
        }


        public DataSet ExcelProcess(string path)
        {
            DataTable dtExcel = new DataTable();
            dtExcel.Columns.Add("item_number");
            dtExcel.Columns.Add("item_value");
            dtExcel.Columns.Add("special_value");
            dtExcel.Columns.Add("order_value");
            dtExcel.Columns.Add("change_value");


  

            using (FileStream fileStream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream))
                {

                    var dataSetConfiguration = new ExcelDataSetConfiguration()
                    {
                        UseColumnDataType = false
                    };
                    
                    // This reads each Sheet into a DataTable and each column is of type System.Object
                   DataSet excelDataSet = excelReader.AsDataSet(dataSetConfiguration);
                    return excelDataSet;
                }
            }



        }
    }
}
