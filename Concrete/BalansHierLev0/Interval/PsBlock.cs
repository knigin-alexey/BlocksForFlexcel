using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Interval
{
    class PsBlock : AbstractBodyFooterBlock<string, InternalData>
    {
        public PsBlock(string data) : base(data)
        {
            AffectToNestingLevelBehavior = new ConcreteAffectToNestingLevelBehavior();
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            StartRow = Parent.GetUsedRows() + 1;
            UsedRows = StartRow;
            StartCol = NestingLevel;
            UsedCols = Parent.GetUsedCols();

            xls.SetCellValue(UsedRows, NestingLevel, Data);
            xls.MergeCells(UsedRows, NestingLevel, UsedRows, UsedCols);
        }

        protected override void RenderFooter(ExtendedXlsFile xls)
        {
            xls.SetRowOutlineLevel(StartRow + 1, UsedRows, 1);

            UsedRows++;
            xls.SetCellValueStyled(UsedRows, NestingLevel, "Итого по " + Data, TFlxFontStyles.Bold);
            xls.MergeCells(UsedRows, NestingLevel, UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);

            RenderSummaryBehavior behav = new RenderSummaryBehavior(this, TFlxFontStyles.Bold);
            behav.RenderSummary(xls, GetInnerFirstEntryFormulas);
            //behav.AddSummaryFormulas(xls);
        }
    }
}
