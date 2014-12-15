using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelOut
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Dictionary<string, ISheet> allSheets = new Dictionary<string, ISheet>();
        
        DataSet ds = new DataSet();

        IWorkbook workbook = null;
        
        private void btnTst_Click(object sender, EventArgs e)
        {
            string v = this.listBox1.SelectedItem.ToString();
            if (allSheets.ContainsKey(v))
            {
                getZuoye(allSheets[v]);
            }
        }


        private string getZuoye(ISheet st)
        {
            int rs=0, re=99;

            for (int i = 0; i < st.LastRowNum; i++)
            {
                if (1 < st.GetRow(i).Cells.Count && st.GetRow(i).Cells[1].ToString().Trim() == "作业")
                {
                    rs = i;
                }
                if (1 < st.GetRow(i).Cells.Count && st.GetRow(i).Cells[1].ToString().Trim() == "课堂表现")
                {
                    re = i;
                }


            }


            for (int j = rs; j < re; j++)
            {
                var con = st.GetRow(j).Cells[2].ToString();
                MessageBox.Show(j.ToString() + "j" + con + st.GetRow(j).Cells[2].CellStyle.GetFont(;
            }
            
            return "";

        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            //e.AllowedEffect = DragDropEffects.All;
            e.Effect = DragDropEffects.Link;
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Effect == DragDropEffects.Link)
            {
                var ss = e.Data.GetFormats(true);
                string[] ff = (string[])e.Data.GetData(DataFormats.FileDrop, true);
                if (System.IO.Directory.Exists(ff[0]))
                {
                    MessageBox.Show("请拖入文件,不要拖入目录");
                }
                else
                {
                    for (int i = 0; i < ff.Length; i++)
                    {
                        loadFile(ff[i]);
                        buildModel();
                        //this.dataGridView1.DataSource = dt;
                    }
                }
            }
        }

        string printDataTable(DataTable dt)
        {
            StringBuilder sb = new StringBuilder();

            foreach (DataColumn item in dt.Columns)
            {
                sb.Append(item.ColumnName.ToString() + "\t");
            }
            sb.AppendLine();

            foreach (DataRow item in dt.Rows)
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sb.Append(item[i].ToString().Trim() + "\t");
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        DataTable ConvertToDataTable(int sheetAt)
        {
            ISheet sheet = workbook.GetSheetAt(sheetAt);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
   
            DataTable dt = new DataTable();
            //for (int j = 0; j < 7; j++)
            //{
            //    dt.Columns.Add(Convert.ToChar(((int)'A') + j).ToString());
            //}

            while (rows.MoveNext())
            {
                IRow row = (IRow)rows.Current;
                DataRow dr = dt.NewRow();

                while (row.LastCellNum > dt.Columns.Count)
                {
                    dt.Columns.Add();
                }

                for (int i = 0; i < row.LastCellNum; i++)
                {
                    ICell cell = row.GetCell(i);
                    
                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        dr[i] = cell.ToString();
                    }
                }
                dt.Rows.Add(dr);
            }
            string classType = getClassType(dt);
            dt.TableName = workbook.GetSheetName(sheetAt) + "."+ classType;
            allSheets.Add(dt.TableName, sheet);
            return dt;
        }

        string getClassType(DataTable dt)
        {
            foreach (DataRow item in dt.Rows)
            {
                if (item[1].ToString().Trim() == "科目")
                {
                    return item[2].ToString().Trim();
                }
            }
            return "未知科目" + new Random().Next(99999).ToString();
        }

        string getClass(DataTable dt)
        {
            foreach (DataRow item in dt.Rows)
            {
                if (item[1].ToString().Trim() == "班级")
                {
                    return item[2].ToString().Trim();
                }
            }
            return "未知班级" + new Random().Next(99999).ToString();
        }

        void loadFile(string filename)
        {
            if (!System.IO.File.Exists(filename))
            {
                throw new IOException("文件没有找到");
            }

            using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                try
                {
                    if (System.IO.Path.GetExtension(filename) == ".xls")
                    {
                        workbook = new HSSFWorkbook(file);
                    }
                    else if (System.IO.Path.GetExtension(filename) == ".xlsx")
                    {
                        workbook = new XSSFWorkbook(file);
                    }
                }
                catch (IOException err)
                {
                    MessageBox.Show("不能识别的文件格式");
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
            }
        }

        private void buildModel()
        {
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ds.Tables.Add(ConvertToDataTable(i));
            }
            this.listBox1.Items.Clear();
            foreach (DataTable dt in ds.Tables)
            {
                this.listBox1.Items.Add(dt.TableName);
            }
        }

        //public DataTable ExcelToDataTable(string sheetName, bool isFirstRowColumn)
        //{
            //ISheet sheet = null;
            //DataTable data = new DataTable();
            //int startRow = 0;
            //try
            //{
            //    fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            //    if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            //        workbook = new XSSFWorkbook(fs);
            //    else if (fileName.IndexOf(".xls") > 0) // 2003版本
            //        workbook = new HSSFWorkbook(fs);

            //    if (sheetName != null)
            //    {
            //        sheet = workbook.GetSheet(sheetName);
            //    }
            //    else
            //    {
            //        sheet = workbook.GetSheetAt(0);
            //    }
            //    if (sheet != null)
            //    {
            //        IRow firstRow = sheet.GetRow(0);
            //        int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

            //        if (isFirstRowColumn)
            //        {
            //            for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
            //            {
            //                DataColumn column = new DataColumn(firstRow.GetCell(i).StringCellValue);
            //                data.Columns.Add(column);
            //            }
            //            startRow = sheet.FirstRowNum + 1;
            //        }
            //        else
            //        {
            //            startRow = sheet.FirstRowNum;
            //        }

            //        //最后一列的标号
            //        int rowCount = sheet.LastRowNum;
            //        for (int i = startRow; i <= rowCount; ++i)
            //        {
            //            IRow row = sheet.GetRow(i);
            //            if (row == null) continue; //没有数据的行默认是null　　　　　　　

            //            DataRow dataRow = data.NewRow();
            //            for (int j = row.FirstCellNum; j < cellCount; ++j)
            //            {
            //                if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
            //                    dataRow[j] = row.GetCell(j).ToString();
            //            }
            //            data.Rows.Add(dataRow);
            //        }
            //    }

            //    return data;
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("Exception: " + ex.Message);
            //    return null;
            //}
        //}

    }
}
