using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    public abstract class AbstractBodyFooterBlock<T, H> : AbstractDataBlock<T, H>
    {
        public AbstractBodyFooterBlock(T data) : base(data)
        {
        }

        public AbstractBodyFooterBlock(T data, H internalData) : base(data, internalData)
        {
        }

        protected abstract void RenderBody(ExtendedXlsFile xls);
        protected abstract void RenderFooter(ExtendedXlsFile xls);

        public override void Render(ExtendedXlsFile xls)
        {
            RenderBody(xls);
            RenderInnerBlocks(xls);
            RenderFooter(xls);
        }
    }
}
