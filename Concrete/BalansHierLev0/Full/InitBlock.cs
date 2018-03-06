using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Data.Full;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Full
{
    public class InitBlock : AbstractBodyFooterBlock<InitData, InternalData>
    {
        public InitBlock(InitData data, InternalData internalData) : base(data, internalData)
        {
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            InternalData.Initialize(GetMaxParentMaxNestingLevel(), Data.VoltageClassPoints);

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
