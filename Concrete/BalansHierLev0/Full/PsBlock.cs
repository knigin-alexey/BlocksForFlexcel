using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using Proryv.AskueARM2.Server.DBAccess.Internal.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Full
{
    /// <summary>
    /// По RenderFooter этого класса четко видно, что должна быть обертка выше, 
    /// которая будет делать вывод значений исходя из значения TExportExcelAdapterType
    /// </summary>
    public class PsBlock : AbstractBodyFooterBlock<string, InternalData>
    {
        public PsBlock(string data, InternalData internalData) : base(data, internalData)
        {
            AffectToNestingLevelBehavior = new ConcreteAffectToNestingLevelBehavior();
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            UsedRows = Parent.GetUsedRows() + 1;
            StartRow = UsedRows;
            StartCol = NestingLevel;

            xls.SetCellValue(UsedRows, NestingLevel, Data);
            xls.MergeCells(UsedRows, NestingLevel, UsedRows, InternalData.TotalColumnsCount);
        }

        protected override void RenderFooter(ExtendedXlsFile xls)
        {
            xls.SetRowOutlineLevel(StartRow + 1, UsedRows, 1);

            UsedRows++;
            xls.SetCellValueStyled(UsedRows, NestingLevel, "Итого по " + Data, TFlxFontStyles.Bold);
            xls.MergeCells(UsedRows, NestingLevel, UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);

            RenderSummaryBehavior behav = new RenderSummaryBehavior(this, TFlxFontStyles.Bold);
            behav.RenderSummary(xls, GetInnerFirstEntryFormulas);
            behav.AddSummaryFormulas(xls, GetInnerFirstEntryFormulas);
        }
    }
}
