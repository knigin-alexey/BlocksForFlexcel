using Proryv.AskueARM2.Both.VisualCompHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.Classes
{
    /// <summary>
    /// Просто класс с данными, которые будут помещаться в List, для отображения
    /// в блоке Title.
    /// </summary>
    public class TitleInfo
    {
        public string Name { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string DateRepresentation
        {
            get
            {
                if (StartDate != DateTime.MinValue && EndDate != DateTime.MinValue)
                {
                    return FlexcelExtensions.GetDateTimePeriodString(StartDate, EndDate);
                }
                else
                {
                    throw new ArgumentException("Для использования этого свойства необходимо задать дату начала и дату окончания");
                }
            }
        }

        public TitleInfo(string name, DateTime startDate, DateTime endDate)
        {
            Name = name;
            StartDate = startDate;
            EndDate = endDate;
        }
    }
}
