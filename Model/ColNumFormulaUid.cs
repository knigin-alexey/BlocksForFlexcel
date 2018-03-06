using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    public class ColNumFormulaUid : Tuple<int, Guid>
    {
        public int ColumnNumber
        {
            get
            {
                return Item1;
            }
        }
        public Guid FormulaUid
        {
            get
            {
                return Item2;
            }
        }

        public ColNumFormulaUid(int col) : base(col, Guid.NewGuid())
        {

        }

        public ColNumFormulaUid(int col, Guid uid) : base(col, uid)
        {

        }
    }
}
