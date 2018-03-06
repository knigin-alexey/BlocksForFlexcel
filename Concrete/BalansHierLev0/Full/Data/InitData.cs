using Proryv.AskueARM2.Server.DBAccess.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Data.Full
{
    public class InitData
    {
        public List<enumVoltageClassPoint> VoltageClassPoints { get; set; }

        public InitData(List<enumVoltageClassPoint> voltageClassPoints)
        {
            VoltageClassPoints = voltageClassPoints;
        }
    }
}
