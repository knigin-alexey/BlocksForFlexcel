using FlexCel.Core;
using Proryv.AskueARM2.Server.DBAccess.Internal.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula
{
    public class FormulaElement
    {
        private int col;
        private int row;
        private double value;
        private EnumExcelFormulaOperators operBefore;

        public FormulaElement(int col, int row, double value, EnumExcelFormulaOperators operBefore)
        {
            this.col = col;
            this.row = row;
            this.value = value;
            this.operBefore = operBefore;
        }

        public int Col
        {
            get
            {
                return col;
            }
        }

        public int Row
        {
            get
            {
                return row;
            }
        }

        public double Value
        {
            get
            {
                return value;
            }
        }

        public EnumExcelFormulaOperators Oper
        {
            get
            {
                return operBefore;
            }
        }

        public string StringRepresentation
        {
            get
            {
                return Oper.Oper() + RowCol;
            }
        }

        public string RowCol
        {
            get
            {
                return TCellAddress.EncodeColumn(Col) + Row.ToString();
            }
        }
    }
}
