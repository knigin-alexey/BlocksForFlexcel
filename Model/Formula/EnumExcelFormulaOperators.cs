using Proryv.AskueARM2.Server.DBAccess.Internal.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula
{
    public enum EnumExcelFormulaOperators
    {
        [ExcelFormulaOperator("+")]
        Plus,
        [ExcelFormulaOperator("-")]
        Minus,
        //[ExcelFormulaOperator("*")]
        //Multiply,
        //[ExcelFormulaOperator("/")]
        //Divide,
        //[ExcelFormulaOperator("SUM")]
        //Sum
    }
}
