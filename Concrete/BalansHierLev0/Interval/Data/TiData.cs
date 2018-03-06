using Proryv.AskueARM2.Server.DBAccess.Internal.TClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Interval.Data
{
    class TiData
    {
        public string Name { get; set; }
        public List<TVALUES_DB> InputInterval { get; set; }
        public List<TVALUES_DB> OutputInterval { get; set; }

        public TiData(string Name)
        {
            this.Name = Name;
            InputInterval = new List<TVALUES_DB>();
            OutputInterval = new List<TVALUES_DB>();
        }
    }
}
