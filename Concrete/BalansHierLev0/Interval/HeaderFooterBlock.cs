using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Interval
{
    class HeaderFooterBlock : AbstractBodyFooterBlock<Data.HeaderFooterData, InternalData>
    {
        public HeaderFooterBlock(Data.HeaderFooterData data) : base(data)
        {
            AffectToNestingLevelBehavior = new ConcreteNotAffectToNestingLevelBehavior();
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            StartRow = Parent.GetUsedRows() + 1;
            UsedRows = StartRow;
            StartCol = InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName];
            UsedCols = InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1;

            xls.SetCellValue(UsedRows + 1,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                "Субъект/Подстанция/Присоединение",
                InternalData.GetHeaderFormat(xls));
            xls.MergeCells(UsedRows + 1,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                UsedRows + 1,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1); //Нужно объединить ячейку до "Прием кВт.ч."

            foreach (var dateTimePeriod in Data.PeriodsList)
            {
                UsedCols++;
                xls.SetCellValue(UsedRows, UsedCols, 
                    dateTimePeriod.DateStart.ToString("dd.MM.yyyy") + " " + dateTimePeriod.DateStart.ToString("HH:mm") + " - " + dateTimePeriod.DateEnd.ToString("HH:mm"), 
                    InternalData.GetHeaderFormat(xls));
                xls.MergeCells(UsedRows, UsedCols, UsedRows, UsedCols + 2);

                xls.SetCellValue(UsedRows + 1, UsedCols, "Прием кВт.ч.", InternalData.GetHeaderFormat(xls));
                xls.SetCellValue(UsedRows + 1, UsedCols + 1, "Отдача кВт.ч.", InternalData.GetHeaderFormat(xls));
                xls.SetCellValue(UsedRows + 1, UsedCols + 2, "Сальдо кВт.ч.", InternalData.GetHeaderFormat(xls));

                UsedCols++;
                UsedCols++;
            }

            xls.SetBorderRangeCells(UsedRows, StartCol, UsedRows, UsedCols, System.Drawing.Color.Black, FlexCel.Core.TFlxBorderStyle.Dashed);
            xls.SetBorderRangeCells(UsedRows + 1, StartCol, UsedRows + 1, UsedCols, System.Drawing.Color.Black, FlexCel.Core.TFlxBorderStyle.Thin);

            UsedRows++;
        }

        protected override void RenderFooter(ExtendedXlsFile xls)
        {
            UsedRows++;

            xls.SetCellValueStyled(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                "Потери в " + ((Data.VoltageClass == enumVoltageClassGlobal.V220AndLower) ? "сети 220 кВ и ниже" : "сети 330 кВ и выше"),
                TFlxFontStyles.Bold);
            xls.MergeCells(UsedRows,
               InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
               UsedRows,
               InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);

            xls.SetCellValueStyled(UsedRows + 1, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                ("Относительные потери в " +
                    ((Data.VoltageClass == enumVoltageClassGlobal.V220AndLower) ? "сети 220 кВ и ниже" : "сети 330 кВ и выше") +
                    " к отпуску в РСК и потребителям"), TFlxFontStyles.Bold);
            xls.MergeCells(UsedRows + 1,
               InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
               UsedRows + 1,
               InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);

            UsedCols = InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1;
            for (int i = 0; i < InternalData.NumbersValues; i++)
            {
                UsedCols++;
                var lossesFormula = GetFirstLevelFormulas(InternalData.IntervalFormulUids[i].SummBalPartSaldoForLosses);
                xls.SetFormula(UsedRows, UsedCols, lossesFormula);

                string relativeLossesFormula;

                var dividerFormula = GetFirstLevelFormulas(InternalData.IntervalFormulUids[i].SummBalPartSaldoForLossesDivider);

                if (dividerFormula.Elements.Count > 0)
                {
                    if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
                    {
                        relativeLossesFormula = "=(" + lossesFormula.StringRepresentation() + ")/("
                            + dividerFormula.StringRepresentation() + ")*100";
                        xls.SetCellFloatValue(UsedRows + 1,
                                UsedCols,
                                new TFormula(relativeLossesFormula));
                    }
                    else
                    {
                        double relativeLosses = lossesFormula.DoubleRepresentation()
                            / dividerFormula.DoubleRepresentation() * 100;
                        xls.SetCellFloatValue(UsedRows + 1, UsedCols, relativeLosses);
                    }
                }
                else
                {
                    xls.SetCellFloatValue(UsedRows + 1, UsedCols, 0.000);
                }

                xls.MergeCells(UsedRows, UsedCols, UsedRows, UsedCols + 2);
                xls.MergeCells(UsedRows + 1, UsedCols, UsedRows + 1, UsedCols + 2);

                UsedCols++;
                UsedCols++;
            }
           
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
                xls.MergeCells(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, UsedRows, UsedCols);
            }
            if (InternalData.FlagOVremark)
            {
                UsedRows++;
                xls.SetCellFloatValue(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], 0.0);
                xls.SetCellFontColor(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], Color.Red);
                xls.SetCellValue(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, " - обходной выключатель");
                xls.MergeCells(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, UsedRows, UsedCols);
            }
            if (InternalData.FlagManualStatusChange)
            {
                UsedRows++;
                xls.SetCellBkColor(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], Color.Yellow);
                xls.SetCellValue(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, " - вручную установлен некоммерческий статус");
                xls.MergeCells(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] + 1, UsedRows, UsedCols);
            }
        }
    }
}
