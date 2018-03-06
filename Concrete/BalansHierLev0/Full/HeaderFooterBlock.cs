using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Data.Full;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using Proryv.AskueARM2.Server.DBAccess.Internal.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Full
{
    public class HeaderFooterBlock : AbstractBodyFooterBlock<HeaderFooterData, InternalData>
    {
        public HeaderFooterBlock(HeaderFooterData data, InternalData header) : base(data, header)
        {
            AffectToNestingLevelBehavior = new ConcreteNotAffectToNestingLevelBehavior();
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            UsedRows = Parent.GetUsedRows() + 1;
            StartRow = UsedRows;
            StartCol = InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName];

            xls.SetCellValue(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                "Субъект/Подстанция/Присоединение",
                InternalData.GetHeaderFormat(xls));
            xls.MergeCells(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                UsedRows + 1,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1); //Нужно объединить ячейку до "Прием кВт.ч."

            xls.SetCellValue(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm],
                "Прием кВт.ч.",
                InternalData.GetHeaderFormat(xls));
            xls.MergeCells(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm],
                UsedRows + 1,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm]);
            if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
            {
                xls.SetColWidth(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm],
                    xls.GetColWidth(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm]) +
                        Convert.ToInt32(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] * ExcelMetrics.ColMult(xls)));
            }

            xls.SetCellValue(UsedRows,
                InternalData.HeaderInputColumnNumbers.Values.Select(x => x.ColumnNumber).Min(),
                "В том числе по уровням напряжения",
                InternalData.GetHeaderFormat(xls));
            if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
            {
                xls.SetColOutlineLevel(InternalData.HeaderInputColumnNumbers.Values.Select(x => x.ColumnNumber).Min(),
                    InternalData.HeaderInputColumnNumbers.Values.Select(x => x.ColumnNumber).Max(),
                    1);
            }
            xls.MergeCells(UsedRows,
                InternalData.HeaderInputColumnNumbers.Values.Select(x => x.ColumnNumber).Min(),
                UsedRows,
                InternalData.HeaderInputColumnNumbers.Values.Select(x => x.ColumnNumber).Max());

            foreach (var voltage in InternalData.HeaderInputColumnNumbers)
            {
                xls.SetCellValue(UsedRows + 1,
                    voltage.Value.ColumnNumber,
                    ((((double)voltage.Key) / 100.0)).ToString() + "кВ",
                    InternalData.GetHeaderFormat(xls));

                if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
                {
                    xls.SetColWidth(voltage.Value.ColumnNumber,
                        xls.GetColWidth(voltage.Value.ColumnNumber) +
                            Convert.ToInt32(voltage.Value.ColumnNumber * ExcelMetrics.ColMult(xls)));
                }
            }

            xls.SetCellValue(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm],
                "Отдача кВт.ч.",
                InternalData.GetHeaderFormat(xls));
            xls.MergeCells(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm],
                UsedRows + 1,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm]);
            if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
            {
                xls.SetColWidth(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm],
                    xls.GetColWidth(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm]) +
                        Convert.ToInt32(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm] * ExcelMetrics.ColMult(xls)));
            }


            xls.SetCellValue(UsedRows,
                InternalData.HeaderOutputColumnNumbers.Values.Select(x => x.ColumnNumber).Min(),
                "В том числе по уровням напряжения",
                InternalData.GetHeaderFormat(xls));
            if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
            {
                xls.SetColOutlineLevel(InternalData.HeaderOutputColumnNumbers.Values.Select(x => x.ColumnNumber).Min(),
                    InternalData.HeaderOutputColumnNumbers.Values.Select(x => x.ColumnNumber).Max(),
                    1);
            }
            xls.MergeCells(UsedRows,
                InternalData.HeaderOutputColumnNumbers.Values.Select(x => x.ColumnNumber).Min(),
                UsedRows,
                InternalData.HeaderOutputColumnNumbers.Values.Select(x => x.ColumnNumber).Max());

            foreach (var voltage in InternalData.HeaderOutputColumnNumbers)
            {
                xls.SetCellValue(UsedRows + 1,
                    voltage.Value.ColumnNumber,
                    ((((double)voltage.Key) / 100.0)).ToString() + "кВ",
                    InternalData.GetHeaderFormat(xls));

                if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
                {
                    xls.SetColWidth(voltage.Value.ColumnNumber,
                        xls.GetColWidth(voltage.Value.ColumnNumber) +
                            Convert.ToInt32(voltage.Value.ColumnNumber * ExcelMetrics.ColMult(xls)));
                }
            }

            xls.SetCellValue(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm],
                "Сальдо кВт.ч.",
                InternalData.GetHeaderFormat(xls));
            xls.MergeCells(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm],
                UsedRows + 1,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm]);
            if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
            {
                xls.SetColWidth(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm],
                    xls.GetColWidth(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm]) +
                        Convert.ToInt32(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm] * ExcelMetrics.ColMult(xls)));
            }

            UsedRows++;
            UsedCols = InternalData.TotalColumnsCount;

            SetBorderAllCellsInBlock(xls, System.Drawing.Color.Black, FlexCel.Core.TFlxBorderStyle.Thin);

            xls.SetColWidth(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName], Convert.ToInt32(30f * ExcelMetrics.ColMult(xls)));
            xls.SetColWidth(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.ItemInLev1], Convert.ToInt32(30f * ExcelMetrics.ColMult(xls)));

            if (InternalData.HeaderColumnNumbers.ContainsKey(HeaderBal0LogicalParts.ItemInLev3))
            {
                xls.SetColWidth(InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.ItemInLev2], Convert.ToInt32(30f * ExcelMetrics.ColMult(xls)));
            }

            if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
            {
                xls.CollapseOutlineCols(1, TCollapseChildrenMode.Collapsed);
            }
        }

        protected override void RenderFooter(ExtendedXlsFile xls)
        {
            UsedRows++;
            xls.SetCellValue(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                "Итого по ЕНЕС, кВт*ч",
                InternalData.GetSectionNameFormat(xls));
            xls.MergeCells(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                UsedRows,
                InternalData.TotalColumnsCount);
            UsedRows++;

            xls.SetCellValueStyled(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                    "Всего по объектам ЕНЭС " + ((Data.VoltageClass == enumVoltageClassGlobal.V220AndLower) ? "220" : "330") + " кВ и выше", TFlxFontStyles.Bold);
            xls.MergeCells(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);

            RenderSummaryBehavior behav = new RenderSummaryBehavior(this, TFlxFontStyles.None);
            behav.RenderSummary(xls, GetFirstLevelFormulas);
            UsedRows++;

            xls.SetCellValueStyled(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                "Потери в " + ((Data.VoltageClass == enumVoltageClassGlobal.V220AndLower) ? "сети 220 кВ и ниже" : "сети 330 кВ и выше"),
                TFlxFontStyles.Bold);
            xls.MergeCells(UsedRows,
               InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
               UsedRows,
               InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);

            var saldoFormula = GetFirstLevelFormulas(FormulaNamesEnum.Balance220330SaldoSummary.GetFormulaUid());
            var balPartInRelLosses = GetFirstLevelFormulas(FormulaNamesEnum.Balance220330RelativeLosses.GetFormulaUid());
            xls.SetFormula(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm],
                saldoFormula);
            UsedRows++;

            xls.SetCellValueStyled(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                ("Относительные потери в " +
                    ((Data.VoltageClass == enumVoltageClassGlobal.V220AndLower) ? "сети 220 кВ и ниже" : "сети 330 кВ и выше") +
                    " к отпуску в РСК и потребителям"), TFlxFontStyles.Bold);
            xls.MergeCells(UsedRows,
               InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
               UsedRows,
               InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);

            string relativeLossesFormula;
            if (balPartInRelLosses.Elements.Count > 0)
            {
                if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
                {
                    relativeLossesFormula = "=(" + saldoFormula.StringRepresentation() + ")/("
                        + balPartInRelLosses.StringRepresentation() + ")*100";
                    xls.SetCellFloatValue(UsedRows,
                            InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm],
                            new TFormula(relativeLossesFormula));
                }
                else
                {
                    double relativeLosses = saldoFormula.DoubleRepresentation()
                        / balPartInRelLosses.DoubleRepresentation() * 100;
                    xls.SetCellFloatValue(UsedRows,
                        InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm],
                        relativeLosses);
                }
            }
            else
            {
                xls.SetCellFloatValue(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], 0.000);
            }

            UsedRows++;

            xls.SetCellValueStyled(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                   "Относительные нормативные потери в сети " +
                   ((Data.VoltageClass == enumVoltageClassGlobal.V220AndLower) ? "220" : "330") + "кВ, %",
                   TFlxFontStyles.Bold);
            xls.MergeCells(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);
            xls.SetCellFloatValue(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], Data.HighLimit);
            UsedRows++;
            UsedRows++;

            if (InternalData.FlNoData || InternalData.FlagOVremark || InternalData.FlagManualStatusChange)
            {
                UsedRows++;
                UsedRows++;
                xls.SetCellValue(UsedRows,
                    InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1,
                    "Примечание : ", InternalData.GetFootnoteFormat(xls));
            }
            if (InternalData.FlNoData)
            {
                UsedRows++;
                xls.SetCellBkColor(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], Color.LightPink);
                xls.SetCellValue(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, " - нет данных по одной или нескольким точкам измерения");
                xls.MergeCells(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, UsedRows, InternalData.TotalColumnsCount);
            }
            if (InternalData.FlagOVremark)
            {
                UsedRows++;
                xls.SetCellFloatValue(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], 0.0);
                xls.SetCellFontColor(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], Color.Red);
                xls.SetCellValue(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, " - обходной выключатель");
                xls.MergeCells(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, UsedRows, InternalData.TotalColumnsCount);
            }
            if (InternalData.FlagManualStatusChange)
            {
                UsedRows++;
                xls.SetCellBkColor(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], Color.Yellow);
                xls.SetCellValue(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, " - вручную установлен некоммерческий статус");
                xls.MergeCells(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, UsedRows, InternalData.TotalColumnsCount);
            }
        }
    }
}
