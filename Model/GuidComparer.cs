using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    public class GuidComparer : IComparer<Guid>
    {
        public int Compare(Guid x, Guid y)
        {
            return x.GetSequenceNum().CompareTo(y.GetSequenceNum());
        }
    }

    public static class GuidExt
    {
        public static Guid Sequence(this Guid g, int sequenceNum)
        {
            var bytes = g.ToByteArray();

            BitConverter.GetBytes(sequenceNum).CopyTo(bytes, 0);

            return new Guid(bytes);
        }

        public static int GetSequenceNum(this Guid g)
        {
            return BitConverter.ToInt32(g.ToByteArray(), 0);
        }
    }
}
