using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Data.Full;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using Proryv.AskueARM2.Server.DBAccess.Internal.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Full
{
    public class BalancePartBlock : AbstractBodyFooterBlock<BalancePartData, InternalData>
    {
        public BalancePartBlock(BalancePartData data, InternalData internalData) : base(data, internalData)
        {
            AffectToNestingLevelBehavior = new ConcreteAffectToNestingLevelBehavior();
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            UsedRows = Parent.GetUsedRows() + 1;
            StartRow = UsedRows;
            StartCol = InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName];
            UsedCols = InternalData.TotalColumnsCount;

            xls.SetCellValue(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                Data.Name,
                InternalData.GetSectionNameFormat(xls));
            xls.MergeCells(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName],
                UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm]);
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
                AddCellToFormula(FormulaNamesEnum.Balance220330RelativeLosses.GetFormulaUid(),
                    InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm],
                    UsedRows,
                    GetInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330SaldoSummary.GetFormulaUid()).DoubleRepresentation(), EnumExcelFormulaOperators.Minus);
            }

            AddCellToFormula(FormulaNamesEnum.Balance220330AllOutputSummary.GetFormulaUid(),
                       InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm],
                       GetUsedRows(),
                       GetInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330OutputSummary.GetFormulaUid()).DoubleRepresentation(),
                       EnumExcelFormulaOperators.Plus);

            AddCellToFormula(FormulaNamesEnum.Balance220330AllSaldoSummary.GetFormulaUid(),
                       InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm],
                       GetUsedRows(),
                       GetInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330SaldoSummary.GetFormulaUid()).DoubleRepresentation(),
                       EnumExcelFormulaOperators.Plus);

            SetBorderAllCellsInBlock(xls, Color.Gray, TFlxBorderStyle.Dotted);
        }
    }
}
