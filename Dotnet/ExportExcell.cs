using System;
using System.IO;
using System.Data;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Web;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace WebApi.BuildHtmlReport
{
    public  class ExportExcell
    {
        public  string DownloadExcelKetXuat(DataTable data, string NameReport, string Fromdate, string Todate)
        {
            int rowStart = 8;
            const int colStart = 0;
            string path = HttpContext.Current.Server.MapPath("~/App_Data/FileExcel/" + NameReport + ".xlsx");
            MemoryStream outFileStream = new MemoryStream();
            var tenBaoCao = string.Empty;
            FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read);
            IWorkbook templateWorkbook = new XSSFWorkbook(fs);
            fs.Close();
            var font = templateWorkbook.CreateFont();
            font.FontName = "Times New Roman";
            font.FontHeightInPoints = 12;
            ISheet sheet = templateWorkbook.GetSheetAt(0);
            IRow row = sheet.GetRow(1);
            row.Cells[1].SetCellValue(Fromdate);
            row.Cells[1].CellStyle.SetFont(font);
            row.Cells[1].CellStyle.Alignment = HorizontalAlignment.Left;

            IRow row1 = sheet.GetRow(2);
            row1.Cells[1].SetCellValue(Todate.ToString());
            row1.Cells[1].CellStyle.SetFont(font);
            row1.Cells[1].CellStyle.Alignment = HorizontalAlignment.Left;

            // Insert dữ liệu của header, body
            int numAllRowHeader = 0;
            ICellStyle styleCellDetail = templateWorkbook.CreateCellStyle();
            rowStart = 5;
            var r = sheet.GetRow(rowStart).Cells[0];
            styleCellDetail.CloneStyleFrom(sheet.GetRow(rowStart).Cells[0].CellStyle);
            BuildBody(data, templateWorkbook, styleCellDetail, rowStart + numAllRowHeader, colStart);

            templateWorkbook.Write(outFileStream);
            byte[] bytes;
            bytes = outFileStream.ToArray();
            string base64 = Convert.ToBase64String(bytes,0,bytes.Length);
            outFileStream.Close();
            return base64;
        }
        public  void BuildBody(DataTable dataBody, IWorkbook templateWorkbook, ICellStyle styleCellDetail, int rowStart, int colStart)
        {
            ISheet sheet = templateWorkbook.GetSheetAt(0);
            var font = templateWorkbook.CreateFont();
            font.FontName = "Arial";
            font.FontHeightInPoints = 9;

            // Tạo style có font : bold
            ICellStyle styleCellBold = templateWorkbook.CreateCellStyle();
            styleCellBold.CloneStyleFrom(styleCellDetail);

            //Số hàng và số cột hiện tại
            int numRowCur = rowStart;
            int numColCur = colStart;

            //Tạo hết tất cả các cell của body
            for (int i = 0; i < dataBody.Rows.Count; i++)
            {
                IRow rowCur = CreateRow(ref sheet, numRowCur, dataBody.Columns.Count);
                numRowCur++;
            }
            // Insert dữ liệu vào các cell body
            numRowCur = rowStart;
            foreach (DataRow item in dataBody.Rows)
            {
                ICellStyle style = styleCellDetail;
                IRow rowCur = sheet.GetRow(numRowCur);
                for (int i = 0; i < dataBody.Columns.Count; i++)
                {
                    if (item[i].GetType().Name == "Decimal")
                    {
                        var value = Convert.ToDouble(Convert.ToDecimal(item[i].ToString().Trim()).ToString("#,0.##################", System.Globalization.CultureInfo.InvariantCulture));
                        rowCur.Cells[numColCur].SetCellValue(value);
                    }
                    else
                    {
                        rowCur.Cells[numColCur].SetCellValue(item[i].ToString());
                    }
                    rowCur.Cells[numColCur].CellStyle = style;
                    numColCur++;
                }
                numColCur = colStart;
                numRowCur++;
            }
        }
        public  IRow CreateRow(ref ISheet worksheet, int numRow, int numCell = 500)
        {
            var row = worksheet.GetRow(numRow);
            if (row == null)
            {
                worksheet.CreateRow(numRow);
                row = worksheet.GetRow(numRow);
            }

            for (int i = 0; i < numCell; i++)
            {
                CreateCell(ref row, i);
            }
            return row;
        }
        public  void CreateCell(ref IRow row, int numCell)
        {
            ICell cell = row.GetCell(numCell);

            if (cell == null)
            {
                row.CreateCell(numCell);
            }
        }
    }
}