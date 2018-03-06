using Proryv.AskueARM2.Server.DBAccess.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Data.Full
{
    public class HeaderFooterData
    {
        public enumVoltageClassGlobal VoltageClass { get; set; }
        public double HighLimit { get; set; }

        public HeaderFooterData(enumVoltageClassGlobal VoltageClass, double HighLimit)
        {
            this.VoltageClass = VoltageClass;
            this.HighLimit = HighLimit;
        }
    }
}
