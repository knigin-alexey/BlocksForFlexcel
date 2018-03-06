using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.Classes;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    /// <summary>
    /// Этому классу блока требуются данные для отображения
    /// </summary>
    public class StubBlock : AbstractBodyBlock<Stub, Stub>
    {
        public StubBlock(Stub data) : base(data)
        {

        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
        }
    }
}
