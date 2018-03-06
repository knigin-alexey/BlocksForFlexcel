using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Interval
{
    class RenderSummaryBehavior
    {
        AbstractBlock<InternalData> block;
        TFlxFontStyles style;

        public RenderSummaryBehavior(AbstractBlock<InternalData> block, TFlxFontStyles style)
        {
            this.block = block;
            this.style = style;
        }

        public void RenderSummary(ExtendedXlsFile xls, Func<Guid, ExcelFormula> getInnerFirstEntryFormulas)
        {
            int currentCol = block.InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1;
            for (int i = 0; i < block.InternalData.NumbersValues; i++)
            {
                currentCol++;
                xls.SetFormulaStyled(block.GetUsedRows(),
                    currentCol,
                    getInnerFirstEntryFormulas(block.InternalData.IntervalFormulUids[i].SummInForPs),
                    style);
                block.AddCellToFormula(block.InternalData.IntervalFormulUids[i].SummInForPs, currentCol, block.GetUsedRows(),
                    getInnerFirstEntryFormulas(block.InternalData.IntervalFormulUids[i].SummInForPs).DoubleRepresentation(),
                    Model.Formula.EnumExcelFormulaOperators.Plus);
                xls.SetFormulaStyled(block.GetUsedRows(),
                    currentCol + 1,
                    getInnerFirstEntryFormulas(block.InternalData.IntervalFormulUids[i].SummOutForPs),
                    style);
                block.AddCellToFormula(block.InternalData.IntervalFormulUids[i].SummOutForPs, currentCol + 1, block.GetUsedRows(),
                    getInnerFirstEntryFormulas(block.InternalData.IntervalFormulUids[i].SummOutForPs).DoubleRepresentation(),
                    Model.Formula.EnumExcelFormulaOperators.Plus);
                xls.SetFormulaStyled(block.GetUsedRows(),
                    currentCol + 2,
                    getInnerFirstEntryFormulas(block.InternalData.IntervalFormulUids[i].SummSaldoForPs),
                    style);
                block.AddCellToFormula(block.InternalData.IntervalFormulUids[i].SummSaldoForPs, currentCol + 2, block.GetUsedRows(),
                    getInnerFirstEntryFormulas(block.InternalData.IntervalFormulUids[i].SummSaldoForPs).DoubleRepresentation(),
                   Model.Formula.EnumExcelFormulaOperators.Plus);

                currentCol++;
                currentCol++;
            }
        }

        public void AddSummaryFormulas(ExtendedXlsFile xls, Func<Guid, ExcelFormula> getInnerFirstEntryFormulas)
        {
            int currentCol = block.InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1;
            for (int i = 0; i < block.InternalData.NumbersValues; i++)
            {
                currentCol++;
                block.AddCellToFormula(block.InternalData.IntervalFormulUids[i].SummBalPartSaldoForLosses, currentCol + 2, block.GetUsedRows(), getInnerFirstEntryFormulas(block.InternalData.IntervalFormulUids[i].SummSaldoForPs).DoubleRepresentation(), Model.Formula.EnumExcelFormulaOperators.Plus);
                currentCol++;
                currentCol++;
            }
        }
    }
}
