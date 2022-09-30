using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = NetOffice.ExcelApi;

namespace acad01
{
    public static partial class ExcelTool
    {

        public static Excel.Worksheet GenerateExcelSheet()
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook book = excelApp.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)book.Worksheets[1];
            return sheet;
        }

        public static void MergeCells(this Excel.Worksheet sheet, int topIndex, int leftIndex, int downIndex, int rightIndex)
        {
            Excel.Range range = sheet.Range(sheet.Cells[topIndex, leftIndex], sheet.Cells[downIndex, rightIndex]);
            range.Merge();
        }

        public static void GenerateCePingTable(PingCeTable tb)
        {
           
            Excel.Worksheet sheet = GenerateExcelSheet();
            //merge样例表头及内容
            sheet.MergeCells(tb.examSample.titleRowNo,
                tb.examSample.sampleStarColoumNo,
                tb.examSample.titleRowNo,
                tb.examSample.sampleStarColoumNo + tb.examSample.sampleColoums);
            sheet.MergeCells(tb.examSample.titleRowNo + 1,
                tb.examSample.sampleStarColoumNo,
                tb.examSample.titleRowNo + tb.examSample.samplRowCount,
                tb.examSample.sampleStarColoumNo + tb.examSample.sampleColoums);
            //merge 考核样例
            sheet.Cells[tb.examSample.titleRowNo, tb.examSample.sampleStarColoumNo].Value = "考核图样";
            sheet.MergeCells(2, 7, 2, 15);
            sheet.Cells[2, 7].Value = "考核信息";
            //todo 

        }

        public static void DrawElements(Element element, Excel.Worksheet sheet, int startRow, int startColoum)
        {
            sheet.Cells[startRow, startColoum].Value = "项目";

            sheet.MergeCells(startRow, startColoum + 1, startRow, startColoum + 6);
            sheet.Cells[startRow, startColoum + 1].Value = "类型";

            int sizedHeaderRow = startRow+1;
            int sizedValueRow = sizedHeaderRow + 1;
            int sizedEndRow = sizedHeaderRow + element.sizedElements.Length;
            sheet.MergeCells(sizedHeaderRow, startColoum, sizedEndRow, startColoum);
            sheet.Cells[sizedHeaderRow, startColoum].Value = "零件尺寸检验";

            string[] sizedHeader = { "尺寸类型", "公称尺寸", "上偏差", "下偏差", "上极限尺寸", "下极限尺寸" };

            for (int i = 0; i < sizedHeader.Length; i++)
            {
                sheet.Cells[sizedHeaderRow, i + startColoum + 1].Value = sizedHeader[i];
            }
            for (int i = 0; i < element.sizedElements.Length; i++)
            {

                sheet.Cells[sizedValueRow + i, startColoum + 1].Value = element.sizedElements[i].sizeType;
                sheet.Cells[sizedValueRow + i, startColoum + 2].Value = element.sizedElements[i].baseSize;
                sheet.Cells[sizedValueRow + i, startColoum + 3].Value = element.sizedElements[i].upperSize;
                sheet.Cells[sizedValueRow + i, startColoum + 4].Value = element.sizedElements[i].lowerSize;
                sheet.Cells[sizedValueRow + i, startColoum + 5].Value = element.sizedElements[i].baseSize + element.sizedElements[i].upperSize;
                sheet.Cells[sizedValueRow + i, startColoum + 6].Value = element.sizedElements[i].baseSize + element.sizedElements[i].lowerSize;

            }

            int gToleranceHeaderRow = sizedEndRow + 1;
            int gToleranceValueRow = gToleranceHeaderRow + 1;
            int gToleranceEndRow = gToleranceHeaderRow + element.geometricalTolerances.Length;

            sheet.MergeCells(gToleranceHeaderRow, startColoum, gToleranceEndRow, startColoum);
            sheet.Cells[gToleranceHeaderRow, startColoum].Value = "形位公差";

            sheet.Cells[gToleranceHeaderRow, startColoum + 1].Value = "公差类型";
            sheet.Cells[gToleranceHeaderRow, startColoum + 2].Value = "公差精度";
            sheet.MergeCells(gToleranceHeaderRow, startColoum + 2, gToleranceHeaderRow, startColoum + 6);

            for (int i = 0; i < element.geometricalTolerances.Length; i++)
            {
                sheet.Cells[gToleranceValueRow + i, startColoum + 1].Value = element.geometricalTolerances[i].ToneranceType;
                sheet.Cells[gToleranceValueRow + i, startColoum + 2].Value = element.geometricalTolerances[i].TonerancePrecision;
                sheet.Cells[gToleranceValueRow + i, startColoum + 1].Font.Name = "gdt"; //设置符号字体为gdt。
                sheet.MergeCells(gToleranceValueRow + i, startColoum + 2, gToleranceValueRow + i, startColoum + 6);
            }

            int sRoughnessHeaderRow = gToleranceEndRow + 1;
            int sRoughnessValueRow = sRoughnessHeaderRow + 1;
            int sRoughnessEndRow = sRoughnessHeaderRow + element.surfaceRoughnesses.Length;

            sheet.MergeCells(sRoughnessHeaderRow, startColoum, sRoughnessEndRow, startColoum);
            sheet.Cells[sRoughnessHeaderRow, startColoum].Value = "表面粗糙度";

            sheet.Cells[sRoughnessHeaderRow, startColoum + 1].Value = "粗糙度类别";
            sheet.Cells[sRoughnessHeaderRow, startColoum + 2].Value = "粗糙度值";
            sheet.MergeCells(sRoughnessHeaderRow, startColoum + 2, sRoughnessHeaderRow, startColoum + 6);

            for (int i = 0; i < element.surfaceRoughnesses.Length; i++)
            {
                sheet.Cells[sRoughnessValueRow + i, startColoum + 1].Value = element.surfaceRoughnesses[i].RoughnessType;
                sheet.Cells[sRoughnessValueRow + i, startColoum + 2].Value = element.surfaceRoughnesses[i].RoughnessValue;
                sheet.MergeCells(sRoughnessValueRow + i, startColoum + 2, sRoughnessValueRow + i, startColoum + 6);
            }

            int otherHeaderRow = sRoughnessEndRow + 1;
            int otherValeRow = otherHeaderRow + 1;
            int otherEndRow = otherHeaderRow + element.otherRequirements.Length;

            sheet.MergeCells(otherHeaderRow, startColoum, otherEndRow, startColoum);
            sheet.Cells[otherHeaderRow, startColoum].Value = "其他要求";
            for (int i = 0; i < element.otherRequirements.Length; i++)
            {
                sheet.Cells[otherValeRow + i, startColoum + 1].Value = element.otherRequirements[i].requirement;
                sheet.MergeCells(otherValeRow + i, startColoum + 1, otherValeRow + i, startColoum + 6);
            }

            int safetyHeaderRow = otherEndRow + 1;
            int safetyValeRow = safetyHeaderRow + 1;
            int safetyEndRow = safetyHeaderRow + element.safetyRequirements.Length;

            sheet.MergeCells(safetyHeaderRow, startColoum, safetyEndRow, startColoum);
            sheet.Cells[safetyHeaderRow, startColoum].Value = "安全要求";
            for (int i = 0; i < element.safetyRequirements.Length; i++)
            {
                sheet.Cells[safetyValeRow + i, startColoum + 1].Value = element.safetyRequirements[i].safetyRequirement;
                sheet.MergeCells(safetyValeRow + i, startColoum + 1, safetyValeRow + i, startColoum + 6);
            }

        }
    }
}
