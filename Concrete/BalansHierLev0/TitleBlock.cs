using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.Classes;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Both.VisualCompHelpers;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0
{
    public class TitleBlock : AbstractBodyBlock<TitleInfo, InternalData>
    {
        public TitleBlock(TitleInfo data) : base(data)
        {
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            StartCol = InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.BalPartName];
            StartRow = 1;
            UsedRows = 1;

            UsedCols = InternalData.TotalColumnsCount;

            xls.SetCellValueStyled(UsedRows, 1,
                Data.Name
                + " \n " +
                Data.DateRepresentation, 
                TFlxFontStyles.Bold);
            if (xls.CurrentExportType != TExportExcelAdapterType.toXLS)
            {
                xls.SetCellAlignH(UsedRows, 1, THFlxAlignment.center);
            }
            xls.MergeCells(UsedRows, 1, UsedRows + 1, UsedCols);
            UsedRows++;
        }
    }
}
