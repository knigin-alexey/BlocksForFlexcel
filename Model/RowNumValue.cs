using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    public class RowNumValue : Tuple<int, double>
    {
        public int RowNumber
        {
            get
            {
                return Item1;
            }
        }
        public double Value
        {
            get
            {
                return Item2;
            }
        }

        public RowNumValue(int row) : base(row, 0)
        {

        }

        public RowNumValue(int row, double value) : base(row, value)
        {

        }
    }
}
