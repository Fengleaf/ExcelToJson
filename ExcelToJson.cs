using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Drawing;

namespace ExcelToJson
{
    public partial class ExcelToJson : Form
    {
        public string JsonPath = "./Json";

        private List<string> filesNames = new List<string>();
        private List<Label> fileLabels = new List<Label>();
        private Label nowSelectedLabel = null;

        public ExcelToJson()
        {
            InitializeComponent();
        }

        private void OnButton1Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Multiselect = true;
                dialog.Filter = "Excel Files(*.xlsx)|*.xlsx|Excel Files(*.xls)|*.xls";
                var isValid = dialog.ShowDialog();

                if (isValid == DialogResult.OK)
                {
                    // 對選擇的所有檔案
                    foreach (string s in dialog.FileNames)
                    {
                        // 取得excel表名稱
                        string[] names = s.Split('\\');
                        string name = names[names.Length - 1];
                        if (filesNames.Contains(s))
                            return;
                        filesNames.Add(s);
                        Label label = new Label();
                        label.Text = name;
                        label.Click += SelectFileLabel;
                        label.Margin = new Padding(0, 0, 0, 6);
                        label.AutoSize = false;
                        label.Width = fileListPanel.Width;
                        label.Dock = DockStyle.Right;
                        fileLabels.Add(label);
                        fileListPanel.Controls.Add(label);
                    }
                }
            }
            catch
            {

            }
        }

        private void SelectFileLabel(object sender, EventArgs e)
        {
            if (nowSelectedLabel != null)
            {
                nowSelectedLabel.BackColor = SystemColors.ActiveBorder;
                nowSelectedLabel.ForeColor = SystemColors.ControlText;
            }
            if (nowSelectedLabel == (Label)sender)
            {
                nowSelectedLabel = null;
                return;
            }
            nowSelectedLabel = (Label)sender;
            nowSelectedLabel.BackColor = Color.Blue;
            nowSelectedLabel.ForeColor = Color.White;
        }

        private void DeleteLabel(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete && nowSelectedLabel != null)
            {
                fileLabels.Remove(nowSelectedLabel);
                fileListPanel.Controls.Remove(nowSelectedLabel);
                Controls.Remove(nowSelectedLabel);
                nowSelectedLabel.Dispose();
                nowSelectedLabel = null;
            }
        }

        private void SingleExportButton_Click(object sender, EventArgs e)
        {
            try
            {
                // 對單一檔案
                int index = fileLabels.IndexOf(nowSelectedLabel);
                if (index < 0)
                {
                    MessageBox.Show("沒有選擇的檔案", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                ExportFile(filesNames[index]);
                MessageBox.Show("讀取完成!!", "成功", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            catch (Exception)
            {
                MessageBox.Show("發生問題!!", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void ExportButton_Click(object sender, EventArgs e)
        {
            try
            {
                // 對選擇的所有檔案
                foreach (string s in filesNames)
                    ExportFile(s);
                MessageBox.Show("讀取完成!!", "成功", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
            catch (Exception ex)
            {
                MessageBox.Show("發生問題!!", "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportFile(string fileName)
        {
            IWorkbook workBook;
            using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
            {
                if (Path.GetExtension(fileName).ToLower() == ".xls")
                    workBook = new HSSFWorkbook(fs);
                else
                    workBook = new XSSFWorkbook(fs);
            }
            // 先讀總表
            ISheet sheet = workBook.GetSheet("總表");
            for (int i = 1; i < sheet.LastRowNum; i++)
            {
                // 子表名
                string subSheetName = sheet.GetRow(i).GetCell(1).StringCellValue;
                // 要轉換的json名字
                string jsonName = sheet.GetRow(i).GetCell(2).StringCellValue + ".json";
                ISheet subSheet = workBook.GetSheet(subSheetName);
                // 資料
                List<Dictionary<string, object>> totalData = new List<Dictionary<string, object>>();
                if (subSheet != null)
                {
                    // 第二個 row 為 key
                    IRow keyRow = subSheet.GetRow(1);
                    // 資料型別的 row
                    IRow dataTypeRow = subSheet.GetRow(2);
                    for (int row = 3; row <= subSheet.LastRowNum; row++)
                    {
                        Dictionary<string, object> data = new Dictionary<string, object>();
                        IRow dataRow = subSheet.GetRow(row);
                        for (int col = 0; col < keyRow.LastCellNum; col++)
                        {
                            string dataType = dataTypeRow.GetCell(col).StringCellValue;
                            ICell cell = dataRow.GetCell(col);
                            if (dataType == "int")
                                data[keyRow.GetCell(col).StringCellValue] = Convert.ToInt32(cell.NumericCellValue);
                            else if (dataType == "str")
                                data[keyRow.GetCell(col).StringCellValue] = cell.StringCellValue;
                        }
                        totalData.Add(data);
                    }
                    using (StreamWriter writer = new StreamWriter(Path.Combine(JsonPath, jsonName)))
                    {
                        writer.WriteLine(JsonConvert.SerializeObject(totalData));
                    }
                }
            }
            workBook.Close();
        }
    }
}
