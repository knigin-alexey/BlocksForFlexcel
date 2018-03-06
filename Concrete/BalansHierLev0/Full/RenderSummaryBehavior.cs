using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using Proryv.AskueARM2.Server.DBAccess.Internal.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Full
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
            if (getInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330InputSummary.GetFormulaUid()).Elements.Count > 0)
            {
                xls.SetFormulaStyled(block.GetUsedRows(),
                    block.InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm],
                    getInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330InputSummary.GetFormulaUid()),
                    style);
            }

            if (getInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330OutputSummary.GetFormulaUid()).Elements.Count > 0)
            {
                xls.SetFormulaStyled(block.GetUsedRows(),
                    block.InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm],
                    getInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330OutputSummary.GetFormulaUid()),
                    style);
            }

            if (getInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330SaldoSummary.GetFormulaUid()).Elements.Count > 0)
            {
                xls.SetFormulaStyled(block.GetUsedRows(),
                    block.InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm],
                    getInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330SaldoSummary.GetFormulaUid()),
                    style);
            }

            //TODO: дублирование
            foreach (var voltage in block.InternalData.HeaderInputColumnNumbers)
            {
                if (getInnerFirstEntryFormulas(voltage.Value.FormulaUid).Elements.Count > 0)
                {
                    xls.SetFormulaStyled(block.GetUsedRows(),
                        voltage.Value.ColumnNumber,
                        getInnerFirstEntryFormulas(voltage.Value.FormulaUid),
                        style);
                }
            }
            //TODO: дублирование
            foreach (var voltage in block.InternalData.HeaderOutputColumnNumbers)
            {
                if (getInnerFirstEntryFormulas(voltage.Value.FormulaUid).Elements.Count > 0)
                {
                    xls.SetFormulaStyled(block.GetUsedRows(),
                    voltage.Value.ColumnNumber,
                    getInnerFirstEntryFormulas(voltage.Value.FormulaUid),
                    style);
                }
            }
        }

        public void AddSummaryFormulas(ExtendedXlsFile xls, Func<Guid, ExcelFormula> getInnerFirstEntryFormulas)
        {
            block.AddCellToFormula(FormulaNamesEnum.Balance220330InputSummary.GetFormulaUid(),
                block.InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm],
                block.GetUsedRows(),
                getInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330InputSummary.GetFormulaUid()).DoubleRepresentation(),
                EnumExcelFormulaOperators.Plus);

            block.AddCellToFormula(FormulaNamesEnum.Balance220330OutputSummary.GetFormulaUid(),
                  block.InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm],
                  block.GetUsedRows(),
                  getInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330OutputSummary.GetFormulaUid()).DoubleRepresentation(),
                  EnumExcelFormulaOperators.Plus);

            block.AddCellToFormula(FormulaNamesEnum.Balance220330SaldoSummary.GetFormulaUid(),
                   block.InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm],
                   block.GetUsedRows(),
                   getInnerFirstEntryFormulas(FormulaNamesEnum.Balance220330SaldoSummary.GetFormulaUid()).DoubleRepresentation(),
                   EnumExcelFormulaOperators.Plus);

            foreach (var voltage in block.InternalData.HeaderInputColumnNumbers)
            {
                if (getInnerFirstEntryFormulas(voltage.Value.FormulaUid).Elements.Count > 0)
                {
                    block.AddCellToFormula(voltage.Value.FormulaUid,
                        voltage.Value.ColumnNumber,
                        block.GetUsedRows(),
                        getInnerFirstEntryFormulas(voltage.Value.FormulaUid).DoubleRepresentation(),
                        EnumExcelFormulaOperators.Plus);
                }
            }

            foreach (var voltage in block.InternalData.HeaderOutputColumnNumbers)
            {
                if (getInnerFirstEntryFormulas(voltage.Value.FormulaUid).Elements.Count > 0)
                {
                    block.AddCellToFormula(voltage.Value.FormulaUid,
                        voltage.Value.ColumnNumber,
                        block.GetUsedRows(),
                        getInnerFirstEntryFormulas(voltage.Value.FormulaUid).DoubleRepresentation(),
                        EnumExcelFormulaOperators.Plus);
                }
            }
        }
    }
}
