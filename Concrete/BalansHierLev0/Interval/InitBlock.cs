using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Interval;
using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.Classes;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Interval
{
    public class InitBlock : AbstractBodyFooterBlock<Stub, InternalData>
    {
        public InitBlock(InternalData internalData) : base(new Stub(), internalData)
        {
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            InternalData.Initialize(GetMaxParentMaxNestingLevel());

            TFlxFormat getDefaultFormat = xls.GetDefaultFormat;
            int formatIndex = xls.DefaultFormatId;
            xls.ActiveSheet = 1;
            if (InternalData.ExportType == TExportExcelAdapterType.toXLS)
            {
                //getDefaultFormat.Font.Name = "Times New Roman";
                getDefaultFormat.Font.Size20 = 188;
            }
            else
            {
                getDefaultFormat.Font.Size20 = 212;
            }
            xls.SetFormat(formatIndex, getDefaultFormat);
        }

        protected override void RenderFooter(ExtendedXlsFile xls)
        {
            if (InternalData.ExportType == TExportExcelAdapterType.toHTML)
            {
                xls.AutofitCols(0.99f);
            }
            else
            {
                xls.AutofitCols(1.28f);
            }
        }
    }
}
