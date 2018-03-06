using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    public interface IAffectToNestingLevelBehavior
    {
        bool IsAffectToNestingLevel();
    }

    public class ConcreteNotAffectToNestingLevelBehavior : IAffectToNestingLevelBehavior
    {
        public bool IsAffectToNestingLevel()
        {
            return false;
        }
    }

    public class ConcreteAffectToNestingLevelBehavior : IAffectToNestingLevelBehavior
    {
        public bool IsAffectToNestingLevel()
        {
            return true;
        }
    }
}
