using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Interval.Data;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using Proryv.AskueARM2.Server.DBAccess.Internal.TClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Interval
{
    class TiBlock : AbstractBodyBlock<Data.TiData, InternalData>
    {
        public TiBlock(Data.TiData data) : base(data)
        {
            AffectToNestingLevelBehavior = new ConcreteAffectToNestingLevelBehavior();
        }

        protected override void RenderBody(ExtendedXlsFile xls)
        {
            StartRow = Parent.GetUsedRows() + 1;
            UsedRows = StartRow;
            StartCol = NestingLevel;

            xls.SetCellValue(UsedRows, StartCol, Data.Name);

            if (StartCol < InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1)
            {
                xls.MergeCells(UsedRows, StartCol, UsedRows, InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1);
            }

            UsedCols = InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm] - 1;

            for (int i = 0; i < InternalData.NumbersValues; i++)
            {
                UsedCols++;
                if (i < Data.InputInterval.Count)
                {
                    xls.SetCellFloatValue(UsedRows, UsedCols, Data.InputInterval[i].F_VALUE);
                    AddCellToFormula(InternalData.IntervalFormulUids[i].SummInForPs, UsedCols, UsedRows, Data.InputInterval[i].F_VALUE, EnumExcelFormulaOperators.Plus);
                    AddCellToFormula(InternalData.IntervalFormulUids[i].SaldoByTi, UsedCols, UsedRows, Data.InputInterval[i].F_VALUE, EnumExcelFormulaOperators.Plus);

                    MarkFlags(Data.InputInterval[i], xls, UsedCols);
                    MarkFlags(Data.InputInterval[i], xls, UsedCols + 2);
                }

                if (i < Data.OutputInterval.Count)
                {
                    xls.SetCellFloatValue(UsedRows, UsedCols + 1, Data.OutputInterval[i].F_VALUE);
                    AddCellToFormula(InternalData.IntervalFormulUids[i].SummOutForPs, UsedCols + 1, UsedRows, Data.OutputInterval[i].F_VALUE, EnumExcelFormulaOperators.Plus);
                    AddCellToFormula(InternalData.IntervalFormulUids[i].SaldoByTi, UsedCols + 1, UsedRows, Data.OutputInterval[i].F_VALUE, EnumExcelFormulaOperators.Minus);
                    MarkFlags(Data.OutputInterval[i], xls, UsedCols + 1);
                    MarkFlags(Data.OutputInterval[i], xls, UsedCols + 2);
                }

                if (i < Data.InputInterval.Count || i < Data.OutputInterval.Count)
                {
                    xls.SetFormula(UsedRows, UsedCols + 2, GetFormula(InternalData.IntervalFormulUids[i].SaldoByTi));
                    AddCellToFormula(InternalData.IntervalFormulUids[i].SummSaldoForPs, UsedCols + 2, UsedRows,
                        GetFormula(InternalData.IntervalFormulUids[i].SaldoByTi).DoubleRepresentation(), EnumExcelFormulaOperators.Plus);
                }

                UsedCols++;
                UsedCols++;
            }
        }

        private void MarkFlags(TVALUES_DB watts, ExtendedXlsFile xls, int col)
        {
            //есть ОВ
            if ((watts.F_FLAG & (VALUES_FLAG_DB.None | VALUES_FLAG_DB.IsOVReplaced)) ==
                                    (VALUES_FLAG_DB.None | VALUES_FLAG_DB.IsOVReplaced))
            {
                xls.SetCellFontColor(UsedRows, col, IsOVReplacedColor);
                xls.SetCellFontColor(UsedRows, col, IsOVReplacedColor);
                InternalData.FlagOVremark = true;
            }
            // нет данных
            if ((watts.F_FLAG & VALUES_FLAG_DB.DataNotFull) == VALUES_FLAG_DB.DataNotFull)
            {
                xls.SetCellBkColor(UsedRows, col, NoDataColor);
                InternalData.FlNoData = true;
            }
            // статус менялся вручную
            if (((watts.F_FLAG & VALUES_FLAG_DB.NotCorrect) == VALUES_FLAG_DB.NotCorrect)
                && ((watts.F_FLAG & VALUES_FLAG_DB.isManualStatusChanged) ==
                 VALUES_FLAG_DB.isManualStatusChanged))
            {
                xls.SetCellBkColor(UsedRows, col, IsManualStatusChangeColor);
                InternalData.FlagManualStatusChange = true;
            }
        }
    }
}
