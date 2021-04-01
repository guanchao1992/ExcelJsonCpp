using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace tablegen2.logic
{
    public static class TableExcelWriter
    {
        public static void genExcel(TableExcelData data, string filePath)
        {
            Util.MakesureFolderExist(Path.GetDirectoryName(filePath));

            var ext = Path.GetExtension(filePath).ToLower();
            if (ext != ".xls" && ext != ".xlsx")
                throw new Exception(string.Format("无法识别的文件扩展名 {0}", ext));

            var workbook = ext == ".xls" ? (IWorkbook)new HSSFWorkbook() : (IWorkbook)new XSSFWorkbook();
            var sheet = workbook.CreateSheet(AppData.Config.SheetNameForData);

            //创建新字体
            var font = workbook.CreateFont();
            font.IsBold = true;

            //创建新样式
            var style = workbook.CreateCellStyle();
            style.SetFont(font);
            style.FillPattern = FillPattern.SolidForeground;
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;

            //默认样式
            for (short i = 0; i < workbook.NumCellStyles; i++)
            {
                workbook.GetCellStyleAt(i).VerticalAlignment = VerticalAlignment.Center;
            }

            try
            {
                IRow row0 = sheet.GetRow(0);
                IRow row1 = sheet.GetRow(1);
                IRow row2 = sheet.GetRow(2);
                if (row0 == null)
                    row0 = sheet.CreateRow(0);
                if (row1 == null)
                    row1 = sheet.CreateRow(1);
                if (row2 == null)
                    row2 = sheet.CreateRow(2);

                row0.Height = 50;

                for (int i = 0; i < data.Headers.Count; i++)
                {
                    var cell0 = row0.GetCell(i);
                    var cell1 = row1.GetCell(i);
                    var cell2 = row2.GetCell(i);
                    if (cell0 == null)
                        cell0 = row0.CreateCell(i);
                    if (cell1 == null)
                        cell1 = row1.CreateCell(i);
                    if (cell2 == null)
                        cell2 = row2.CreateCell(i);

                    cell0.SetCellValue(data.Headers[i].FieldDesc);
                    cell1.SetCellValue(data.Headers[i].FieldType);
                    cell2.SetCellValue(data.Headers[i].FieldName);
                }
            }
            catch (Exception)
            {
                throw new Exception(string.Format("{0} 表创建失败！", filePath));
            }
            var tmppath = Path.Combine(Path.GetDirectoryName(filePath),
                string.Format("{0}.tmp{1}", Path.GetFileNameWithoutExtension(filePath), ext));
            using (var fs = File.Open(tmppath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                workbook.Write(fs);
            }
            var content = File.ReadAllBytes(tmppath);
            File.Delete(tmppath);
            File.WriteAllBytes(filePath, content);
        }
    }
}
