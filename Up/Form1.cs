using Microsoft.Win32;
using Newtonsoft.Json;
using OfficeOpenXml;
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

namespace Up
{
    public partial class Form1 : Form
    {
        int ruNameIdx = 1;
        int ruDeptIdx = 1;
        int ruOdrIdx = 1;
        int ruFacIdx = 1;
        string ruWorkSheet = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = $"{ProductName} {ProductVersion}";
            openFileDialog1.Filter = "Excel File (.xlsx)|*.xlsx";
            textBox1.Text = GetSetting("File1");
            textBox2.Text = GetSetting("File2");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = textBox1.Text;

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                SetSetting("File1", textBox1.Text);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = textBox2.Text;
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                checkedListBox1.Items.Clear();
                label3.Text = "";
                Application.DoEvents();
                textBox2.Text = openFileDialog1.FileName;
                SetSetting("File2", textBox2.Text);
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("真的要退出程式嗎？", "退出程式", MessageBoxButtons.OKCancel) == DialogResult.Cancel)
            {
                e.Cancel = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count > 0)
            {
                Dictionary<string, bool> checkList = new Dictionary<string, bool>();
                foreach (var obj in checkedListBox1.CheckedItems)
                {
                    checkList.Add($"{obj}", false);
                    //obj.

                }

                Console.WriteLine(string.Join(",", checkList));

                var ruExcelFile = new FileInfo(textBox2.Text);
                if (ruExcelFile.Exists)
                {
                    using (var ruExcel = new ExcelPackage(ruExcelFile))
                    {
                        ExcelWorksheet ruSheet = ruExcel.Workbook.Worksheets[ruWorkSheet];
                        var excelFile = new FileInfo(textBox1.Text);

                        using (var excel = new ExcelPackage(excelFile))
                        {
                            var sheet = excel.Workbook.Worksheets[0];
                            //MessageBox.Show($"{sheet.Cells[1, 1].Value}");
                            //foreach (var sheet in excel.Workbook.Worksheets)
                            //{
                            //    MessageBox.Show($"{sheet.Name}");
                            //}
                            //ExcelWorksheet sheet1 = excel.Workbook.Worksheets["MySheet"];

                            var distOdrIdx = 1;
                            var disNameIdx = 1;
                            var distDeptIdx = 1;
                            var distFacIdx = 1;

                            Dictionary<string, int> ruDic = new Dictionary<string, int>();
                            Dictionary<string, int> distDic = new Dictionary<string, int>();

                            while (sheet.Cells[5, distOdrIdx].Value == null || sheet.Cells[5, distOdrIdx].Value.ToString().IndexOf("ORDER NO.") == -1)
                            {
                                distOdrIdx++;
                            }
                            distOdrIdx++;

                            while (sheet.Cells[5, distOdrIdx].Value != null && !string.IsNullOrWhiteSpace(sheet.Cells[5, distOdrIdx].Value.ToString()))
                            {
                                var myType = sheet.Cells[5, distOdrIdx].Value.ToString();
                                if (!distDic.ContainsKey(myType))
                                {
                                    distDic.Add(myType, distOdrIdx);
                                }
                                distOdrIdx++;
                            }

                            while (ruSheet.Cells[5, ruOdrIdx].Value != null && !string.IsNullOrWhiteSpace(ruSheet.Cells[5, ruOdrIdx].Value.ToString()))
                            {
                                var myType = ruSheet.Cells[5, ruOdrIdx].Value.ToString();
                                if (!ruDic.ContainsKey(myType))
                                {
                                    ruDic.Add(myType, ruOdrIdx);
                                }
                                ruOdrIdx++;
                            }

                            while (sheet.Cells[6, disNameIdx].Value == null || sheet.Cells[6, disNameIdx].Value.ToString() != "製單人")
                            {
                                disNameIdx++;
                            }

                            while (sheet.Cells[6, distDeptIdx].Value == null || sheet.Cells[6, distDeptIdx].Value.ToString() != "部門")
                            {
                                distDeptIdx++;
                            }

                            while (sheet.Cells[6, distFacIdx].Value == null || sheet.Cells[6, distFacIdx].Value.ToString() != "廠別")
                            {
                                distFacIdx++;
                            }

                            Console.WriteLine(JsonConvert.SerializeObject(ruDic));
                            /*while (sheet.Cells[6, ruDeptIdx].Value == null || sheet.Cells[6, ruDeptIdx].Value.ToString() != "部門")
                            {
                                ruDeptIdx++;
                            }*/
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show($"請選擇製作人", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            var excelFile = new FileInfo(textBox2.Text);
            if (excelFile.Exists)
            {
                checkedListBox1.Items.Clear();
                Application.DoEvents();
                using (var excel = new ExcelPackage(excelFile))
                {
                    //MessageBox.Show(string.Join(",", excel.Workbook.Worksheets.Select(x => x.Name)));
                    ruWorkSheet = "";
                    foreach (var sheet in excel.Workbook.Worksheets)
                    {
                        if (sheet.Name.StartsWith("WK", StringComparison.OrdinalIgnoreCase))
                        {
                            ruWorkSheet = sheet.Name;
                        }
                    }

                    label3.Text = ruWorkSheet;
                    Application.DoEvents();
                    ExcelWorksheet sheet1 = excel.Workbook.Worksheets[ruWorkSheet];

                    //"UA1J"
                    if (sheet1 != null)
                    {
                        int idx = 7;
                        int RU_May_Idx = 0;
                        ruNameIdx = 1;
                        ruDeptIdx = 1;
                        ruFacIdx = 1;
                        ruOdrIdx = 1;
                        int ruACCOUNTIdx = 1;
                        while (sheet1.Cells[6, ruNameIdx].Value == null || sheet1.Cells[6, ruNameIdx].Value.ToString() != "製單人")
                        {
                            ruNameIdx++;
                        }

                        while (sheet1.Cells[6, ruDeptIdx].Value == null || sheet1.Cells[6, ruDeptIdx].Value.ToString() != "部門")
                        {
                            ruDeptIdx++;
                        }

                        while (sheet1.Cells[5, ruACCOUNTIdx].Value == null || sheet1.Cells[5, ruACCOUNTIdx].Value.ToString() != "ACCOUNT")
                        {
                            ruACCOUNTIdx++;
                        }

                        while (ruFacIdx < sheet1.Dimension.Columns && (sheet1.Cells[6, ruFacIdx].Value == null || sheet1.Cells[6, ruFacIdx].Value.ToString() != "廠別"))
                        {
                            ruFacIdx++;
                        }

                        while (sheet1.Cells[5, ruOdrIdx].Value == null || sheet1.Cells[5, ruOdrIdx].Value.ToString().Trim().IndexOf("ORDER NO.") == -1)
                        {
                            ruOdrIdx++;
                        }
                        ruOdrIdx++;

                        for (RU_May_Idx = sheet1.Dimension.Rows; RU_May_Idx >= 7; RU_May_Idx--)
                        {
                            if (sheet1.Cells[RU_May_Idx, 5].Value != null && sheet1.Cells[RU_May_Idx, 5].Value.ToString().Trim().IndexOf("RU May") != -1)
                            {
                                break;
                            }
                        }

                        Dictionary<string, Dictionary<string, int>> DicRMFac = new Dictionary<string, Dictionary<string, int>>();
                        Dictionary<string, Dictionary<string, int>> DicOrderFac = new Dictionary<string, Dictionary<string, int>>();
                        Dictionary<string, int> DicMayOrder = new Dictionary<string, int>();
                        Dictionary<string, bool> DicOrderNoExistsMay = new Dictionary<string, bool>();
                        StringBuilder compareErr = new StringBuilder();

                        for (var i = 7; i <= sheet1.Dimension.Columns; i++)
                        {
                            if (sheet1.Cells[5, i].Value != null && sheet1.Cells[RU_May_Idx, i].Value != null && int.TryParse(sheet1.Cells[RU_May_Idx, i].Value.ToString(), out int q) && q > 0)
                            {
                                string orderKey = sheet1.Cells[5, i].Value.ToString();
                                if (!DicMayOrder.ContainsKey(orderKey))
                                {
                                    DicMayOrder.Add(orderKey, q);
                                }
                            }
                        }
                        Console.WriteLine(JsonConvert.SerializeObject(DicMayOrder));
                        for (idx = 7; idx <= sheet1.Dimension.Rows; idx++)
                        {
                            if (string.Equals(sheet1.Cells[idx, ruDeptIdx].Value?.ToString(), "UA1J", StringComparison.OrdinalIgnoreCase))
                            {
                                var factory = $"{sheet1.Cells[idx, ruFacIdx].Value}";

                                if (string.Equals(sheet1.Cells[idx, ruACCOUNTIdx].Value?.ToString(), "RM", StringComparison.OrdinalIgnoreCase))
                                {
                                    if (sheet1.Cells[idx, ruFacIdx].Value != null)
                                    {
                                        for (var i = ruOdrIdx; i <= sheet1.Dimension.Columns; i++)
                                        {
                                            if (sheet1.Cells[5, i].Value != null && sheet1.Cells[idx, i].Value != null && int.TryParse(sheet1.Cells[idx, i].Value.ToString(), out int q) && q > 0)
                                            {
                                                string key = sheet1.Cells[5, i].Value.ToString();
                                                if (DicMayOrder.ContainsKey(key))
                                                {
                                                    if (!DicRMFac.ContainsKey(key))
                                                    {
                                                        DicRMFac.Add(key, new Dictionary<string, int>());
                                                    }

                                                    var f = DicRMFac[key];

                                                    if (!f.ContainsKey(factory))
                                                    {
                                                        f.Add(factory, 0);
                                                    }

                                                    f[factory] += q;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (sheet1.Cells[idx, ruFacIdx].Value != null)
                                    {
                                        for (var i = ruOdrIdx; i <= sheet1.Dimension.Columns; i++)
                                        {
                                            if (sheet1.Cells[5, i].Value != null && sheet1.Cells[idx, i].Value != null && int.TryParse(sheet1.Cells[idx, i].Value.ToString(), out int q) && q > 0)
                                            {
                                                string key = sheet1.Cells[5, i].Value.ToString();
                                                if (DicMayOrder.ContainsKey(key))
                                                {
                                                    if (!DicOrderFac.ContainsKey(key))
                                                    {
                                                        DicOrderFac.Add(key, new Dictionary<string, int>());
                                                    }

                                                    var f = DicOrderFac[key];

                                                    if (!f.ContainsKey(factory))
                                                    {
                                                        f.Add(factory, 0);
                                                    }

                                                    f[factory] += q;
                                                }
                                                else if (!DicOrderNoExistsMay.ContainsKey(key))
                                                {
                                                    DicOrderNoExistsMay.Add(key, false);
                                                    compareErr.AppendLine($"\"{key}\" ,不存在於May");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        Console.WriteLine(JsonConvert.SerializeObject(DicRMFac));
                        Console.WriteLine(JsonConvert.SerializeObject(DicOrderFac));
                        //Dictionary<string,>
                        if (compareErr.Length > 0)
                        {
                            compareErr.AppendLine();
                        }

                        foreach (var may in DicMayOrder)
                        {
                            if (DicOrderFac.ContainsKey(may.Key))
                            {
                                int ot = 0;
                                bool hasQtErr = false;
                                foreach (var o in DicOrderFac[may.Key])
                                {
                                    ot += o.Value;
                                }

                                if (ot != may.Value)
                                {
                                    hasQtErr = true;
                                }

                                bool hasFacErr = false;
                                foreach (var o in DicOrderFac[may.Key])
                                {
                                    var oo = DicRMFac[may.Key];
                                    if (!oo.ContainsKey(o.Key) || oo[o.Key] < o.Value)
                                    {
                                        hasFacErr = true;
                                    }
                                }

                                if (hasFacErr || hasQtErr)
                                {
                                    compareErr.AppendLine(may.Key);

                                    if (hasQtErr)
                                        compareErr.AppendLine($@"數量錯誤, sum:{ot:#,##0}不等於May:{may.Value:#,##0}");

                                    if (hasFacErr)
                                    {
                                        if (DicRMFac.ContainsKey(may.Key))
                                            compareErr.AppendLine($"廠區錯誤, sum:{JsonConvert.SerializeObject(DicOrderFac[may.Key])}不等於RM:{JsonConvert.SerializeObject(DicRMFac[may.Key])}");
                                        else
                                            compareErr.AppendLine($"廠區錯誤, sum:{JsonConvert.SerializeObject(DicOrderFac[may.Key])}不存在於RM");
                                    }
                                }
                            }
                            else
                            {
                                compareErr.AppendLine($"{may.Key}\r\n ,May:{may.Value:#,##0}不存在於訂單");
                            }

                            //if (DicOrderFac.ContainsKey(rm.Key))
                            //{
                            //    var o = DicOrderFac[rm.Key];
                            //    foreach (var r in rm.Value)
                            //    {
                            //        if (o.ContainsKey(r.Key))
                            //        {
                            //            if (o[r.Key] != r.Value)
                            //                compareErr.AppendLine($"'{r.Key}'RM:{r.Value:#,##0}不等於Sum:{o[r.Key]:#,##0}");
                            //        }
                            //        else
                            //        {
                            //            compareErr.AppendLine($"ORDER NO'{r.Key}'不存在於RM");
                            //        }
                            //    }
                            //}
                            //else
                            //{
                            //    compareErr.AppendLine($"廠別'{rm.Key}'不存在");
                            //}
                        }

                        if (compareErr.ToString().Length > 0)
                        {
                            var errDialog = new Form2();
                            errDialog.SetError(compareErr.ToString());
                            errDialog.ShowDialog();
                            return;
                        }

                        Dictionary<string, bool> DicName = new Dictionary<string, bool>();
                        idx = 7;
                        while (idx <= sheet1.Dimension.Rows)
                        {
                            if (string.Equals(sheet1.Cells[idx, ruDeptIdx].Value?.ToString(), "UA1J", StringComparison.OrdinalIgnoreCase))
                            {
                                if (sheet1.Cells[idx, ruNameIdx].Value != null)
                                {
                                    var account = $"{sheet1.Cells[idx, ruNameIdx].Value}";

                                    if (!DicName.ContainsKey(account))
                                    {
                                        DicName.Add(account, false);
                                    }
                                }
                            }
                            idx++;
                        }
                        var nameList = DicName.Select(x => x.Key);
                        //MessageBox.Show($"{string.Join(",", nameList)}");

                        checkedListBox1.Items.AddRange(nameList.ToArray());
                        /*if (checkedListBox1.Items.Count > 0)
                            ruSheet = sheet1;*/
                    }
                    else
                    {
                        MessageBox.Show($"找不到'WK*'", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show($"找不到'{excelFile.Name}'", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetSetting(string key)
        {
            var regKeyAppRoot = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\GigaByteUp\Helper");
            if (regKeyAppRoot == null)
            {
                regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\GigaByteUp\Helper");
            }

            var value = regKeyAppRoot.GetValue(key)?.ToString();

            return value;
        }

        private void SetSetting(string key, string value)
        {
            try
            {
                var regKeyAppRoot = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\GigaByteUp\Helper", true);
                if (regKeyAppRoot == null)
                {
                    regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\GigaByteUp\Helper");
                }

                regKeyAppRoot.SetValue(key, value);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex}");
            }
        }
    }
}
