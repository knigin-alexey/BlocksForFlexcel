using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Data.Full;
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
    public class TiBlock : AbstractBodyBlock<TiData, InternalData>
    {
        public TiBlock(TiData data, InternalData innerData) : base(data, innerData)
        {
            AffectToNestingLevelBehavior = new ConcreteAffectToNestingLevelBehavior();
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            UsedRows = Parent.GetUsedRows() + 1;
            StartRow = UsedRows;
            StartCol = NestingLevel;

            xls.SetCellValue(UsedRows, StartCol, Data.Name);

            if (StartCol < InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1)
            {
                xls.MergeCells(UsedRows, StartCol, UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);
            }

            #region Input
            foreach (var inputByVolt in Data.InputByVoltages)
            {
                double inputValue = inputByVolt.Value.ConvertDataUnitDigitFromWatt(EnumUnitDigit.Kilo).F_VALUE;
                xls.SetCellFloatValue(UsedRows,
                    InternalData.HeaderInputColumnNumbers[inputByVolt.Key].ColumnNumber,
                    inputValue);

                //надо добавить в 2 формулы 
                //в input
                AddCellToFormula(HeaderBal0LogicalParts.InputSumm.GetFormulaUid(),
                    InternalData.HeaderInputColumnNumbers[inputByVolt.Key].ColumnNumber, UsedRows, inputValue, EnumExcelFormulaOperators.Plus);
                //по подстанции 
                AddCellToFormula(InternalData.HeaderInputColumnNumbers[inputByVolt.Key].FormulaUid,
                    InternalData.HeaderInputColumnNumbers[inputByVolt.Key].ColumnNumber, UsedRows, inputValue, EnumExcelFormulaOperators.Plus);

                MarkFlags(inputByVolt, xls, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], InternalData.HeaderInputColumnNumbers);
            }
            
            var inputFormula = GetFormula(HeaderBal0LogicalParts.InputSumm.GetFormulaUid());
            if (inputFormula.Elements.Count > 0)
            {
                xls.SetFormula(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm],
                    inputFormula);
                AddCellToFormula(FormulaNamesEnum.Balance220330InputSummary.GetFormulaUid(),
                    InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], UsedRows,
                    inputFormula.DoubleRepresentation(), EnumExcelFormulaOperators.Plus);
                AddCellToFormula(HeaderBal0LogicalParts.SaldoSumm.GetFormulaUid(),
                    InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], UsedRows,
                    inputFormula.DoubleRepresentation(), EnumExcelFormulaOperators.Plus);
            }
            #endregion Input

            #region Output
            foreach (var outputByVolt in Data.OutputByVoltages)
            {
                double outputValue = outputByVolt.Value.ConvertDataUnitDigitFromWatt(EnumUnitDigit.Kilo).F_VALUE;
                xls.SetCellFloatValue(UsedRows,
                    InternalData.HeaderOutputColumnNumbers[outputByVolt.Key].ColumnNumber,
                    outputValue);

                //надо добавить в 2 формулы 
                //в output
                AddCellToFormula(HeaderBal0LogicalParts.OutputSumm.GetFormulaUid(),
                    InternalData.HeaderOutputColumnNumbers[outputByVolt.Key].ColumnNumber,
                    UsedRows,
                    outputValue, EnumExcelFormulaOperators.Plus);
                //по подстанции 
                AddCellToFormula(InternalData.HeaderOutputColumnNumbers[outputByVolt.Key].FormulaUid,
                    InternalData.HeaderOutputColumnNumbers[outputByVolt.Key].ColumnNumber,
                    UsedRows, outputValue, EnumExcelFormulaOperators.Plus);

                MarkFlags(outputByVolt, xls, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm], InternalData.HeaderOutputColumnNumbers);
            }
            
            var outputFormula = GetFormula(HeaderBal0LogicalParts.OutputSumm.GetFormulaUid());
            if (outputFormula.Elements.Count > 0)
            {
                xls.SetFormula(UsedRows,
                InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm],
                outputFormula);
                AddCellToFormula(FormulaNamesEnum.Balance220330OutputSummary.GetFormulaUid(),
                    InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm],
                    UsedRows, outputFormula.DoubleRepresentation(), EnumExcelFormulaOperators.Plus);
                AddCellToFormula(HeaderBal0LogicalParts.SaldoSumm.GetFormulaUid(),
                    InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.OutputSumm], UsedRows,
                    outputFormula.DoubleRepresentation(), EnumExcelFormulaOperators.Minus);
            }
            #endregion Output

            #region Saldo
            var saldoFormula = GetFormula(HeaderBal0LogicalParts.SaldoSumm.GetFormulaUid());

            if (saldoFormula.Elements.Count > 0)
            {
                xls.SetFormula(UsedRows,
                        InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm],
                        saldoFormula);
                AddCellToFormula(FormulaNamesEnum.Balance220330SaldoSummary.GetFormulaUid(),
                    InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm], UsedRows,
                    saldoFormula.DoubleRepresentation(), EnumExcelFormulaOperators.Plus);
            }
            #endregion Saldo
        }

        private void MarkFlags(KeyValuePair<enumVoltageClassPoint, Server.DBAccess.Internal.TClasses.TVALUES_DB> wattsByVolt,
            ExtendedXlsFile xls,
            int groupColNum,
            Dictionary<enumVoltageClassPoint, ColNumFormulaUid> colDict)
        {
            //есть ОВ
            if ((wattsByVolt.Value.F_FLAG & (VALUES_FLAG_DB.None | VALUES_FLAG_DB.IsOVReplaced)) ==
                                    (VALUES_FLAG_DB.None | VALUES_FLAG_DB.IsOVReplaced))
            {
                xls.SetCellFontColor(UsedRows, colDict[wattsByVolt.Key].ColumnNumber, IsOVReplacedColor);
                xls.SetCellFontColor(UsedRows, groupColNum, IsOVReplacedColor);
                xls.SetCellFontColor(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm], IsOVReplacedColor);
                //Data.FlagOVremark = true;
                InternalData.FlagOVremark = true;
            }
            // нет данных
            if ((wattsByVolt.Value.F_FLAG & VALUES_FLAG_DB.DataNotFull) == VALUES_FLAG_DB.DataNotFull)
            {
                xls.SetCellBkColor(UsedRows, colDict[wattsByVolt.Key].ColumnNumber, NoDataColor);
                xls.SetCellBkColor(UsedRows, groupColNum, NoDataColor);
                xls.SetCellBkColor(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm], NoDataColor);
                //Data.IsNoDataStatus = true;
                InternalData.FlNoData = true;
            }
            // статус менялся вручную
            if (((wattsByVolt.Value.F_FLAG & VALUES_FLAG_DB.NotCorrect) == VALUES_FLAG_DB.NotCorrect)
                &&
                ((wattsByVolt.Value.F_FLAG & VALUES_FLAG_DB.isManualStatusChanged) ==
                 VALUES_FLAG_DB.isManualStatusChanged))
            {
                xls.SetCellBkColor(UsedRows, colDict[wattsByVolt.Key].ColumnNumber, IsManualStatusChangeColor);
                xls.SetCellBkColor(UsedRows, groupColNum, IsManualStatusChangeColor);
                xls.SetCellBkColor(UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.SaldoSumm], IsManualStatusChangeColor);
                //Data.FlagManualStatusChange = true;
                InternalData.FlagManualStatusChange = true;
            }
        }
    }
}
