using Proryv.AskueARM2.Server.DBAccess.Internal;
using Proryv.AskueARM2.Server.DBAccess.Internal.TClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Data.Full
{
    public class TiData
    {
        public string Name { get; set; }
        //public bool FlagOVremark { get; set; }
        //public bool IsNoDataStatus { get; set; }
        //public bool FlagManualStatusChange { get; set; }

        public Dictionary<enumVoltageClassPoint, TVALUES_DB> InputByVoltages { get; set; }
        public Dictionary<enumVoltageClassPoint, TVALUES_DB> OutputByVoltages { get; set; }

        public TiData(string Name)
        {
            this.Name = Name;
            InputByVoltages = new Dictionary<enumVoltageClassPoint, TVALUES_DB>();
            OutputByVoltages = new Dictionary<enumVoltageClassPoint, TVALUES_DB>();
        }
    }
}
