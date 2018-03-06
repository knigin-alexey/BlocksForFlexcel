using FlexCel.Core;
using Proryv.AskueARM2.Server.DBAccess.Internal.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula
{
    /// <summary>
    /// Пока будут элементы только одного уровня
    /// </summary>
    public class ExcelFormula
    {
        private List<FormulaElement> elements;

        public List<FormulaElement> Elements
        {
            get
            {
                return elements;
            }
        }

        public ExcelFormula()
        {
            elements = new List<FormulaElement>();
        }

        public void AddElement(FormulaElement element)
        {
            elements.Add(element);
        }

        public void AddElementsRange(List<FormulaElement> elements)
        {
            this.elements.AddRange(elements);
        }

        public string StringRepresentation()
        {
            StringBuilder resultString = new StringBuilder();

            resultString.Append(ConvertElementsListToString(elements));

            return resultString.ToString();
        }

        public double DoubleRepresentation()
        {
            return ConvertElementListToDouble(elements);
        }

        private double ConvertElementListToDouble(List<FormulaElement> formulaList)
        {
            double result = 0;
            foreach (var element in formulaList)
            {
                //Очень сложно сделать адекватный механизм обработки формул с учетом приоритета
                if (element.Oper == EnumExcelFormulaOperators.Plus)
                {
                    result += element.Value;
                }
                else
                {
                    result -= element.Value;
                }
            }
            return result;
        }

        private string ConvertElementsListToString(List<FormulaElement> formulaList)
        {
            StringBuilder resultString = new StringBuilder();

            //все элементы положительные, колонка одна и та же и строки идут подряд, тогда SUM(...)
            List<FormulaElement> continuousList = new List<FormulaElement>();

            var columnsDict = formulaList
                .GroupBy(x => x.Col)
                .ToDictionary(g => g.Key, g => g.ToList());

            foreach (var column in columnsDict)
            {
                var orderedList = column.Value.OrderBy(x => x.Row);
                var elemIter = orderedList.First();
                continuousList.Add(elemIter);

                orderedList.Skip(1)
                    .Aggregate(elemIter, (prev, next) =>
                    {
                        if (next.Row - prev.Row == 1 && next.Oper == prev.Oper)
                        {
                            continuousList.Add(next);
                        }
                        else
                        {
                            resultString.Append(AnalyzeContinuousElementList(continuousList));

                            continuousList.Clear();
                            continuousList.Add(next);
                        }
                        return next;
                    }
                );
            }

            resultString.Append(AnalyzeContinuousElementList(continuousList));

            return resultString.ToString();
        }

        private string AnalyzeContinuousElementList(List<FormulaElement> formulaList)
        {
            string result;
            var colCount = formulaList.GroupBy(x => x.Col)
                    .ToDictionary(g => g.Key, g => g.ToList()).Count;

            if (formulaList.Count > 1 && colCount == 1)
            {
                var orderedContList = formulaList.OrderBy(x => x.Row);

                result = orderedContList.First().Oper.Oper()
                    + "SUM(" + orderedContList.First().RowCol
                    + ":" + orderedContList.Last().RowCol + ")";
            }
            else
            {
                result = string.Join(String.Empty, formulaList.Select(x => x.StringRepresentation));
            }

            return result;
        }
    }
}
