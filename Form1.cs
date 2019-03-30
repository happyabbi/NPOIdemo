using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace NPOIdemo
{
    public partial class Form1 : Form
    {
        private DataTable ProductInfo;

        public Form1()
        {
            InitializeComponent();
        }

        private void createDataTable()
        {
            ProductInfo = new DataTable();

            ProductInfo.Clear();
            ProductInfo.Columns.Add("_Product");
            ProductInfo.Columns.Add("_WH");
            ProductInfo.Columns.Add("_Qty");

            var ProductRow = ProductInfo.NewRow();
            ProductRow["_Product"] = "AS133";
            ProductRow["_WH"] = "W101";
            ProductRow["_Qty"] = "10";
            ProductInfo.Rows.Add(ProductRow);

            ProductRow = ProductInfo.NewRow();
            ProductRow["_Product"] = "AS133";
            ProductRow["_WH"] = "W102";
            ProductRow["_Qty"] = "7";
            ProductInfo.Rows.Add(ProductRow);

            ProductRow = ProductInfo.NewRow();
            ProductRow["_Product"] = "AS133";
            ProductRow["_WH"] = "W103";
            ProductRow["_Qty"] = "5";
            ProductInfo.Rows.Add(ProductRow);

            ProductRow = ProductInfo.NewRow();
            ProductRow["_Product"] = "AS156";
            ProductRow["_WH"] = "W101";
            ProductRow["_Qty"] = "6";
            ProductInfo.Rows.Add(ProductRow);

            ProductRow = ProductInfo.NewRow();
            ProductRow["_Product"] = "TS156";
            ProductRow["_WH"] = "W101";
            ProductRow["_Qty"] = "8";
            ProductInfo.Rows.Add(ProductRow);

            ProductRow = ProductInfo.NewRow();
            ProductRow["_Product"] = "TS156";
            ProductRow["_WH"] = "W102";
            ProductRow["_Qty"] = "8";
            ProductInfo.Rows.Add(ProductRow);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            createDataTable();

            var strFilePath = "template.xlt";
            HSSFWorkbook workbook;
            using (var fs = new FileStream(strFilePath, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = new HSSFWorkbook(fs);
                fs.Close();
            }

            if (workbook != null)
            {
                //load template
                var RawData = (HSSFSheet) workbook.GetSheet("RawData");

                //data to newSheet 
                var hst = RawData;

                for (var i = 0; i < ProductInfo.Rows.Count; i++)
                {
                    var hr = (HSSFRow) hst.CreateRow(i + 1);
                    for (var j = 0; j < ProductInfo.Columns.Count; j++)
                    {
                        var hc = (HSSFCell) hr.CreateCell(j);
                        //Notice!!! Qty is Int,other String.
                        if (ProductInfo.Columns[j].Caption == "_Qty")
                        {
                            hc.SetCellType(CellType.Numeric);
                            if (!string.IsNullOrEmpty(ProductInfo.Rows[i][j].ToString()))
                            {
                                var number = Convert.ToInt32(ProductInfo.Rows[i][j].ToString());
                                hc.SetCellValue(number);
                            }
                            else
                            {
                                hc.SetCellValue(ProductInfo.Rows[i][j].ToString());
                            }
                        }
                        else
                        {
                            hc.SetCellValue(ProductInfo.Rows[i][j].ToString());
                        }
                    }
                }

                //export new EXCEL file
                var filename = "P_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xls";
                var fsExcelNew = new FileStream(filename, FileMode.Create);
                workbook.Write(fsExcelNew);

                //remove sheet
                workbook.RemoveSheetAt(0);
                workbook = null;

                fsExcelNew.Close();

                //Open excel
                //System.Windows.Forms.Application.StartupPath + "\\" + filename;
                var file = @"C:\Windows\explorer.exe";
                Process.Start(file, filename);
            }
        }
    }
}