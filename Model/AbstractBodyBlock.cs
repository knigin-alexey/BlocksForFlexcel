using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    public abstract class AbstractBodyBlock<T, H> : AbstractDataBlock<T, H>
    {
        public AbstractBodyBlock(T data) : base(data)
        {
        }

        public AbstractBodyBlock(T data, H innerData) : base(data, innerData)
        {
        }

        protected abstract void RenderBody(ExtendedXlsFile xls);

        public override void Render(ExtendedXlsFile xls)
        {
            RenderBody(xls);
            RenderInnerBlocks(xls);
        }
    }
}
