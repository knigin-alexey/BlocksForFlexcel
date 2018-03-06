using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.Classes
{
    public class DataForSummary
    {
        public int Col { get; set; }
        public int Row { get; set; }
        public double Summary { get; set; }

        public DataForSummary()
        {

        }

        public DataForSummary(int col, double summary)
        {
            Col = col;
            Summary = summary;
        }
    }
}
