using Proryv.AskueARM2.Server.DBAccess.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Interval.Data
{
    class HeaderFooterData
    {
        public enumVoltageClassGlobal VoltageClass { get; set; }

        public List<DateTimePeriod> PeriodsList { get; set; } 
        public HeaderFooterData(enumVoltageClassGlobal VoltageClass, List<DateTimePeriod> periodsList) //double HighLimit
        {
            this.VoltageClass = VoltageClass;
            this.PeriodsList = periodsList;
            //this.HighLimit = HighLimit;
        }
    }
}
