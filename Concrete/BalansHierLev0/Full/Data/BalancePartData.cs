using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Data.Full
{
    public class BalancePartData
    {
        public string Name { get; set; }
        public bool IsGroupUseInGeneralBalance { get; set; }
        public bool IsGroupUseInRelative { get; set; }

        public BalancePartData(string Name, bool IsGroupUseInGeneralBalance, bool IsGroupUseInRelative)
        {
            this.Name = Name;
            this.IsGroupUseInGeneralBalance = IsGroupUseInGeneralBalance;
            this.IsGroupUseInRelative = IsGroupUseInRelative;
        }
    }
}
