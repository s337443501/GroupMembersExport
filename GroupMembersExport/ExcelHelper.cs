using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace GroupMembersExport
{
    public static class ExcelHelper
    {
        #region
        public static void ExportExcel(DataGridView dgv, string fileName = "", string fontname = "微软雅黑", short fontsize = 10)
        {
            //检测是否有数据
            if (dgv.RowCount <= 0) return;
            //创建主要对象
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = (HSSFSheet)workbook.CreateSheet("Weight");
            //设置字体，大小，对齐方式
            HSSFCellStyle style = (HSSFCellStyle)workbook.CreateCellStyle();
            HSSFFont font = (HSSFFont)workbook.CreateFont();
            font.FontName = fontname;
            font.FontHeightInPoints = fontsize;
            style.SetFont(font);
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //居中对齐
            style.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            style.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            //首行填充黄色
            ICellStyle headerStyle = workbook.CreateCellStyle();
            headerStyle.FillForegroundColor = IndexedColors.Yellow.Index;
            headerStyle.FillPattern = FillPattern.SolidForeground;
            headerStyle.SetFont(font);
            headerStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center; //居中对齐
            headerStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            headerStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            headerStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            headerStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            //添加表头
            HSSFRow dataRow = (HSSFRow)sheet.CreateRow(0);
            for (int i = 0; i < dgv.Columns.Count; i++)
            {
                dataRow.CreateCell(i).SetCellValue(dgv.Columns[i].HeaderText);
                dataRow.GetCell(i).CellStyle = headerStyle;
            }
            //注释的这行是设置筛选的
            //sheet.SetAutoFilter(new CellRangeAddress(0, dgv.Columns.Count, 0, dgv.Columns.Count));
            //添加列及内容
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                sheet.AutoSizeColumn(i);//先来个常规自适应
                dataRow = (HSSFRow)sheet.CreateRow(i + 1);
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    string ValueType = dgv.Rows[i].Cells[j].Value.GetType().ToString();
                    string Value = dgv.Rows[i].Cells[j].Value.ToString();
                    switch (ValueType)
                    {
                        case "System.String"://字符串类型
                            dataRow.CreateCell(j).SetCellValue(Value);
                            break;
                        case "System.DateTime"://日期类型
                            System.DateTime dateV;
                            System.DateTime.TryParse(Value, out dateV);
                            dataRow.CreateCell(j).SetCellValue(dateV);
                            break;
                        case "System.Boolean"://布尔型
                            bool boolV = false;
                            bool.TryParse(Value, out boolV);
                            dataRow.CreateCell(j).SetCellValue(boolV);
                            break;
                        case "System.Int16"://整型
                        case "System.Int32":
                        case "System.Int64":
                        case "System.Byte":
                            int intV = 0;
                            int.TryParse(Value, out intV);
                            dataRow.CreateCell(j).SetCellValue(intV);
                            break;
                        case "System.Decimal"://浮点型
                        case "System.Double":
                            double doubV = 0;
                            double.TryParse(Value, out doubV);
                            dataRow.CreateCell(j).SetCellValue(doubV);
                            break;
                        case "System.DBNull"://空值处理
                            dataRow.CreateCell(j).SetCellValue("");
                            break;
                        default:
                            dataRow.CreateCell(j).SetCellValue("");
                            break;
                    }
                    dataRow.GetCell(j).CellStyle = style;
                    //设置宽度
                    sheet.SetColumnWidth(j, (Value.Length + 10) * 300);
                    //首行冻结
                    sheet.CreateFreezePane(dgv.Columns.Count, 1);
                }
            }
            //保存文件
            string saveFileName = "";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xls";
            saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.FileName = fileName;
            MemoryStream ms = new MemoryStream();
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                saveFileName = saveDialog.FileName;
                if (!CheckFiles(saveFileName))
                {
                    MessageBox.Show("文件占用，请关闭文件后再继续操作！ " + saveFileName);
                    workbook = null;
                    ms.Close();
                    ms.Dispose();
                    return;
                }
                workbook.Write(ms);
                FileStream file = new FileStream(saveFileName, FileMode.Create);
                workbook.Write(file);
                file.Close();
                workbook = null;
                ms.Close();
                ms.Dispose();
                MessageBox.Show("导出完成！", "", MessageBoxButtons.OK);
            }
            else
            {
                workbook = null;
                ms.Close();
                ms.Dispose();
            }
        }
        #endregion

        #region 检测文件被占用 
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);
        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);
        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public static readonly IntPtr HFILE_ERROR = new IntPtr(-1);
        /// <summary>
        /// 检测文件被占用 
        /// </summary>
        /// <param name="FileNames">要检测的文件路径</param>
        /// <returns></returns>
        public static bool CheckFiles(string FileNames)
        {
            if (!File.Exists(FileNames))
            {
                //文件不存在
                return true;
            }
            IntPtr vHandle = _lopen(FileNames, OF_READWRITE | OF_SHARE_DENY_NONE);
            if (vHandle == HFILE_ERROR)
            {
                //文件被占用
                return false;
            }
            //文件没被占用
            CloseHandle(vHandle);
            return true;
        }
        #endregion

    }
}
