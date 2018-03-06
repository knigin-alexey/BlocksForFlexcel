using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Interval
{
    class BalancePartBlock : AbstractBodyFooterBlock<BalansHierLev0.Data.Full.BalancePartData, InternalData>
    {
        public BalancePartBlock(BalansHierLev0.Data.Full.BalancePartData data) : base(data)
        {
            AffectToNestingLevelBehavior = new ConcreteAffectToNestingLevelBehavior();
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            StartRow= Parent.GetUsedRows() + 1;
            UsedRows = StartRow;
            StartCol = InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName];
            UsedCols = Parent.GetUsedCols();

            xls.SetCellValue(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                Data.Name, InternalData.GetSectionNameFormat(xls));
            xls.MergeCells(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                UsedRows, UsedCols);
        }

        protected override void RenderFooter(ExtendedXlsFile xls)
        {
            UsedRows++;
            xls.SetCellValueStyled(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName], "Всего по " + Data.Name, TFlxFontStyles.Bold);
            xls.MergeCells(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);

            RenderSummaryBehavior behav = new RenderSummaryBehavior(this, TFlxFontStyles.Bold);
            behav.RenderSummary(xls, GetInnerFirstEntryFormulas);
            if (Data.IsGroupUseInGeneralBalance)
            {
                behav.AddSummaryFormulas(xls, GetInnerFirstEntryFormulas);
            }

            if (Data.IsGroupUseInRelative)
            {
                UsedCols = InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1;
                for (int i = 0; i < InternalData.NumbersValues; i++)
                {
                    UsedCols++;
                    AddCellToFormula(InternalData.IntervalFormulUids[i].SummBalPartSaldoForLossesDivider,
                        UsedCols + 2, GetUsedRows(), GetInnerFirstEntryFormulas(InternalData.IntervalFormulUids[i].SummSaldoForPs).DoubleRepresentation(), Model.Formula.EnumExcelFormulaOperators.Minus);
                    UsedCols++;
                    UsedCols++;
                }
            }

            SetBorderAllCellsInBlock(xls, Color.Gray, TFlxBorderStyle.Dotted);
        }
    }
}
