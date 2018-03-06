using FlexCel.Core;
using FlexCel.XlsAdapter;
using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.Classes;
using Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    public enum SideToBorder
    {
        Top,
        Right,
        Bottom,
        Left,
        All
    }

    public class ExtendedXlsFile : XlsFile
    {
        public TExportExcelAdapterType CurrentExportType { get; set; }

        public List<DataForSummary> listDataSummary { get; set; }

        public ExtendedXlsFile(TExportExcelAdapterType ExportType, int sheetCount = 1) : base(true)
        {
            this.CurrentExportType = ExportType;
            //this.ListFormulsByID = new Dictionary<Guid, SortedList>();
            listDataSummary = new List<DataForSummary>();

            NewFile(sheetCount);
            TFlxFormat getDefaultFormat = GetDefaultFormat;
            int defaultFormatId = DefaultFormatId;
            //getDefaultFormat.Font.Size20 -= FontOffset;
            SetFormat(defaultFormatId, getDefaultFormat);
            PrintOptions &= ~(TPrintOptions.NoPls | TPrintOptions.Orientation);
            PrintScale = 100;
            PrintNumberOfHorizontalPages = 1;
            PrintNumberOfVerticalPages = 20;
            SetPrintMargins(new TXlsMargins(0.2, 0.2, 0.2, 0.2, 0.2, 0.2));

            if (ExportType == TExportExcelAdapterType.toHTML)
            {
                getDefaultFormat.FillPattern.Pattern = TFlxPatternStyle.Solid;
                getDefaultFormat.FillPattern.FgColor = Color.FromArgb(140, 0xc0, 0xe9);
                SetFormat(defaultFormatId, getDefaultFormat);
            }
        }



        public int SetCellAlignH(int Row, int Col, THFlxAlignment Ha)
        {
            int cellFormat = this.GetCellFormat(Row, Col);
            TFlxFormat format = this.GetFormat(cellFormat);
            format.HAlignment = Ha;
            cellFormat = this.AddFormat(format);
            this.SetCellFormat(Row, Col, cellFormat);
            return cellFormat;
        }

        public int SetCellAlignV(int Row, int Col, TVFlxAlignment Va)
        {
            int cellFormat = this.GetCellFormat(Row, Col);
            TFlxFormat format = this.GetFormat(cellFormat);
            format.VAlignment = Va;
            cellFormat = this.AddFormat(format);
            this.SetCellFormat(Row, Col, cellFormat);
            return cellFormat;
        }

        public int SetCellFontStyle(int Row, int Col, TFlxFontStyles fs)
        {
            int cellFormat = this.GetCellFormat(Row, Col);
            TFlxFormat format = this.GetFormat(cellFormat);
            format.Font.Style = fs;
            cellFormat = this.AddFormat(format);
            this.SetCellFormat(Row, Col, cellFormat);
            return cellFormat;
        }

        public void SetCellValueStyled(int row, int col, object value, TFlxFontStyles fs)
        {
            SetCellValue(row, col, value);
            SetCellFontStyle(row, col, fs);
        }

        private int SetCellFontStyle(int Row, int Col, TFlxFormat Fx)
        {
            int cellFormat = GetCellFormat(Row, Col);
            cellFormat = AddFormat(Fx);
            SetCellFormat(Row, Col, cellFormat);
            return cellFormat;
        }

        public void SetCellValueStyled(int StartRow, int StartCol, TFlxFormat Fx, string Val)
        {
            SetCellValue(StartRow, StartCol, Val);
            SetCellFontStyle(StartRow, StartCol, Fx);
        }

        public void SetCellFloatValue(int Row, int Col, object value)
        {
            SetCellFloatValue(Row, Col, value, 3, true);
        }

        public void SetCellFloatValue(int Row, int Col, object value, int floatDecimals)
        {
            SetCellFloatValue(Row, Col, value, floatDecimals, true);
        }

        public void SetCellFloatValue(int Row, int Col, object value, int floatDecimals, TFlxFontStyles fs)
        {
            SetCellFloatValue(Row, Col, value, floatDecimals);
            SetCellFontStyle(Row, Col, fs);
        }

        private void SetCellFloatValue(int Row, int Col, object value, int FloatDecimals, bool IsNeed0)
        {
            if (Col <= 255)
            {
                if (FloatDecimals < 0)
                {
                    FloatDecimals = 0;
                }
                SetCellValue(Row, Col, value);
                string str = string.Empty;
                for (int i = 1; i <= FloatDecimals; i++)
                {
                    str = str + (IsNeed0 ? "0" : "#");
                }
                string str2 = "# ### ### ##0";
                if (str != string.Empty)
                {
                    str2 = str2 + "." + str;
                }

                if (CurrentExportType == TExportExcelAdapterType.toHTML)
                {
                    if (value != null)
                    {
                        SetCellValue(Row, Col, ((double)value).ToString(str2));
                        SetCellAlignH(Row, Col, THFlxAlignment.right);
                        SetCellAlignV(Row, Col, TVFlxAlignment.center);
                    }
                }
                else
                {
                    int cellFormat = GetCellFormat(Row, Col);
                    TFlxFormat format = GetFormat(cellFormat);
                    format.Format = str2;
                    cellFormat = AddFormat(format);
                    SetCellFormat(Row, Col, cellFormat);
                }
            }
        }

        public void SetFormula(int row, int col, ExcelFormula formula)
        {
            SetFormula(row, col, formula, 3);
        }

        public void SetFormula(int row, int col, ExcelFormula formula, int floatDecimals)
        {
            if (formula != null && formula.Elements.Count > 0)
            {
                if (CurrentExportType == TExportExcelAdapterType.toXLS && formula.StringRepresentation().Length < 1024)
                {
                    SetCellFloatValue(row, col, new TFormula(formula.StringRepresentation().Insert(0, "=")), floatDecimals);
                }
                else
                {
                    SetCellFloatValue(row, col, formula.DoubleRepresentation(), floatDecimals);
                }
            }
        }

        public void SetFormulaStyled(int row, int col, ExcelFormula formula, TFlxFontStyles fs, int floatDecimals = 3)
        {
            SetFormula(row, col, formula, floatDecimals);
            SetCellFontStyle(row, col, fs);
        }

        public void AutofitCols(float adjustment)
        {
            for (int i = 1; i <= ColCount; i++)
            {
                AutofitCol(i, false, adjustment);
                SetColWidth(i, GetColWidth(i) + Convert.ToInt32((10f * ExcelMetrics.ColMult(this))));
            }
        }

        public void SetBorderRangeCells(int startRow, int startCol, int stopRow, int stopCol,
            Color borderColor, TFlxBorderStyle linestyle)
        {
            for (var i = startCol; i <= stopCol; i++)
            {
                for (var j = startRow; j <= stopRow; j++)
                {
                    SetBorderCell(j, i, borderColor, linestyle);
                }
            }
        }

        public void SetBorderCell(int row, int col, Color borderColor,
            TFlxBorderStyle linestyle)
        {
            if (col > 255)
            {
                col = 255;
            }

            var cellFormat = GetCellFormat(row, col);
            var format = GetFormat(cellFormat);
            format.Borders.Bottom.Color = borderColor;
            format.Borders.Left.Color = borderColor;
            format.Borders.Right.Color = borderColor;
            format.Borders.Top.Color = borderColor;
            format.Borders.Bottom.Style = linestyle;
            format.Borders.Left.Style = linestyle;
            format.Borders.Right.Style = linestyle;
            format.Borders.Top.Style = linestyle;
            cellFormat = AddFormat(format);
            SetCellFormat(row, col, cellFormat);
        }

        public void SetCellRangeBorderBySide(int startRow, int startCol, int stopRow, int stopCol,
            Color borderColor, TFlxBorderStyle linestyle, SideToBorder side)
        {
            switch (side)
            {
                case (SideToBorder.Top):
                    //обходим все колонки и рисуем границу в верхней части первой строки
                    for (var i = startCol; i <= stopCol; i++)
                    {
                        SetBorderCellsTop(startRow, i, borderColor, linestyle);
                    }
                    break;
                case (SideToBorder.Right):
                    //обходим все строки и рисуем границу в правой части последней колонки
                    for (var i = startRow; i <= stopRow; i++)
                    {
                        SetBorderCellsRight(i, stopCol, borderColor, linestyle);
                    }
                    break;
                case (SideToBorder.Bottom):
                    for (var i = startCol; i <= stopCol; i++)
                    {
                        SetBorderCellsBottom(stopRow, i, borderColor, linestyle);
                    }
                    break;
                case (SideToBorder.Left):
                    for (var i = startRow; i <= stopRow; i++)
                    {
                        SetBorderCellsLeft(i, startCol, borderColor, linestyle);
                    }
                    break;
                case (SideToBorder.All):
                    for (var i = startCol; i <= stopCol; i++)
                    {
                        SetBorderCellsTop(startRow, i, borderColor, linestyle);
                        SetBorderCellsBottom(stopRow, i, borderColor, linestyle);
                    }
                    for (var i = startRow; i <= stopRow; i++)
                    {
                        SetBorderCellsRight(i, stopCol, borderColor, linestyle);
                        SetBorderCellsLeft(i, startCol, borderColor, linestyle);
                    }
                    break;
                default:
                    break;
            }
        }

        public void SetBorderCellsTop(int row, int col, Color borderColor,
            TFlxBorderStyle linestyle)
        {
            if (col > 255)
            {
                col = 255;
            }

            var cellFormat = GetCellFormat(row, col);
            var format = GetFormat(cellFormat);
            format.Borders.Top.Color = borderColor;
            format.Borders.Top.Style = linestyle;
            cellFormat = AddFormat(format);
            SetCellFormat(row, col, cellFormat);
        }

        public void SetBorderCellsRight(int row, int col, Color borderColor,
            TFlxBorderStyle linestyle)
        {
            if (col > 255)
            {
                col = 255;
            }

            var cellFormat = GetCellFormat(row, col);
            var format = GetFormat(cellFormat);
            format.Borders.Right.Color = borderColor;
            format.Borders.Right.Style = linestyle;
            cellFormat = AddFormat(format);
            SetCellFormat(row, col, cellFormat);
        }

        public void SetBorderCellsBottom(int row, int col, Color borderColor,
            TFlxBorderStyle linestyle)
        {
            if (col > 255)
            {
                col = 255;
            }

            var cellFormat = GetCellFormat(row, col);
            var format = GetFormat(cellFormat);
            format.Borders.Bottom.Color = borderColor;
            format.Borders.Bottom.Style = linestyle;
            cellFormat = AddFormat(format);
            SetCellFormat(row, col, cellFormat);
        }

        public void SetBorderCellsLeft(int row, int col, Color borderColor,
            TFlxBorderStyle linestyle)
        {
            if (col > 255)
            {
                col = 255;
            }

            var cellFormat = GetCellFormat(row, col);
            var format = GetFormat(cellFormat);
            format.Borders.Left.Color = borderColor;
            format.Borders.Left.Style = linestyle;
            cellFormat = AddFormat(format);
            SetCellFormat(row, col, cellFormat);
        }

        public int SetCellBkColor(int Row, int Col, Color bk)
        {
            int cellFormat = GetCellFormat(Row, Col);
            TFlxFormat format = GetFormat(cellFormat);
            format.FillPattern.Pattern = TFlxPatternStyle.Solid;
            format.FillPattern.FgColor = bk;

            cellFormat = AddFormat(format);
            SetCellFormat(Row, Col, cellFormat);
            return cellFormat;
        }

        public void SetRangeBkColor(int rowStart, int rowEnd, int colStart, int colEnd, Color bk)
        {
            for (int r = rowStart; r <= rowEnd; r++)
            {
                for (int c = colStart; c <= colEnd; c++)
                {
                    SetCellBkColor(r, c, bk);
                }
            }
        }

        public int SetCellFontColor(int Row, int Col, Color fc)
        {
            int cellFormat = GetCellFormat(Row, Col);
            TFlxFormat format = GetFormat(cellFormat);
            format.Font.Color = fc;
            cellFormat = AddFormat(format);
            SetCellFormat(Row, Col, cellFormat);
            return cellFormat;
        }
    }
}
