using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using Proryv.AskueARM2.Server.DBAccess.Internal.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0
{
    public enum HeaderBal0LogicalParts
    {
        [Description("Субъект/Подстанция/Присоединение")]
        BalPartName,
        ItemInLev1,
        ItemInLev2,
        [Description("Субъект/Подстанция/Присоединение")]
        ItemInLev3,
        [Description("Прием кВт.ч.")]
        [ExcelFormulaMapper("89896375-0722-422b-8d19-fe87af0419eb")]
        InputSumm,
        [Description("Отдача кВт.ч.")]
        [ExcelFormulaMapper("9dca70a1-c983-4409-8c8a-aaa9ef95e652")]
        OutputSumm,
        [Description("Сальдо   кВт.ч.")]
        [ExcelFormulaMapper("f9cecc8d-162e-4d75-b357-4bdcb28d8d05")]
        SaldoSumm
    }

    public class InternalData
    {
        private bool isInit;
        private Dictionary<HeaderBal0LogicalParts, int> headerColumnNumbers;
        public Dictionary<HeaderBal0LogicalParts, int> HeaderColumnNumbers
        {
            get
            {
                if (!isInit)
                {
                    throw new Exception("Необходимо выполнить метод Initialize");
                }
                return headerColumnNumbers;
            }
        }
        private Dictionary<enumVoltageClassPoint, ColNumFormulaUid> headerInputColumnNumbers;
        /// <summary>
        /// К сожалению, для формул не получится сделать фиксированные гуиды, потому что
        /// Input будет смешиваться с Output
        /// </summary>
        public Dictionary<enumVoltageClassPoint, ColNumFormulaUid> HeaderInputColumnNumbers
        {
            get
            {
                if (!isInit)
                {
                    throw new Exception("Необходимо выполнить метод Initialize");
                }
                return headerInputColumnNumbers;
            }
        }
        private Dictionary<enumVoltageClassPoint, ColNumFormulaUid> headerOutputColumnNumbers;
        /// <summary>
        /// К сожалению, для формул не получится сделать фиксированные гуиды, потому что
        /// Input будет смешиваться с Output
        /// </summary>
        public Dictionary<enumVoltageClassPoint, ColNumFormulaUid> HeaderOutputColumnNumbers
        {
            get
            {
                if (!isInit)
                {
                    throw new Exception("Необходимо выполнить метод Initialize");
                }
                return headerOutputColumnNumbers;
            }
        }

        private int numbersValues;
        public int NumbersValues
        {
            get
            {
                return numbersValues;
            }
        }

        public TExportExcelAdapterType ExportType { get; set; }

        private int totalColumnsCount;
        public int TotalColumnsCount
        {
            get
            {
                return totalColumnsCount;
            }
        }

        public InternalData(TExportExcelAdapterType exportType)
        {
            ExportType = exportType;
        }

        public InternalData(TExportExcelAdapterType exportType, int numbersValues) : this(exportType)
        {
            int i = 0;
            this.numbersValues = numbersValues;
            intervalFormulUids = new Dictionary<int, FormulaIntervalUids>();
            while (i < numbersValues)
            {
                intervalFormulUids.Add(i, new FormulaIntervalUids());
                i++;
            }
        }

        public bool FlNoData { get; set; }
        public bool FlagOVremark { get; set; }
        public bool FlagManualStatusChange { get; set; }

        Dictionary<int, FormulaIntervalUids> intervalFormulUids;
        internal Dictionary<int, FormulaIntervalUids> IntervalFormulUids
        {
            get
            {
                return intervalFormulUids;
            }
        }

        public void Initialize(int nesting)
        {
            headerColumnNumbers = new Dictionary<HeaderBal0LogicalParts, int>();
            headerInputColumnNumbers = new Dictionary<enumVoltageClassPoint, ColNumFormulaUid>();
            headerOutputColumnNumbers = new Dictionary<enumVoltageClassPoint, ColNumFormulaUid>();

            isInit = true;
            totalColumnsCount = 1;
            if (nesting > 4)
            {
                throw new Exception("Что-то много вложенности, не предусмотренно архитектурой. Нужно добавить вложенность в Bal0LogicalParts");
            }

            //походу, не очень нужный кусок кода, т.к. непонятно, как использовать HeaderBal0LogicalParts.ItemInLev1
            //наверное, достаточно будет просто по nesting увеличить totalColumnsCount
            for (int i = 0; i < nesting; i++)
            {
                //эти if'ы сделаны на случай, если кто-то сменит порядок в Bal0LogicalParts
                if (!headerColumnNumbers.ContainsKey(HeaderBal0LogicalParts.BalPartName))
                {
                    headerColumnNumbers.Add(HeaderBal0LogicalParts.BalPartName, totalColumnsCount++);
                }
                else if (!headerColumnNumbers.ContainsKey(HeaderBal0LogicalParts.ItemInLev1))
                {
                    headerColumnNumbers.Add(HeaderBal0LogicalParts.ItemInLev1, totalColumnsCount++);
                }
                else if (!headerColumnNumbers.ContainsKey(HeaderBal0LogicalParts.ItemInLev2))
                {
                    headerColumnNumbers.Add(HeaderBal0LogicalParts.ItemInLev2, totalColumnsCount++);
                }
                else if (!headerColumnNumbers.ContainsKey(HeaderBal0LogicalParts.ItemInLev3))
                {
                    headerColumnNumbers.Add(HeaderBal0LogicalParts.ItemInLev3, totalColumnsCount++);
                }
            }

            headerColumnNumbers.Add(HeaderBal0LogicalParts.InputSumm, totalColumnsCount++);
        }

        public void Initialize(int nesting, List<enumVoltageClassPoint> voltageClassPoints)
        {
            Initialize(nesting);

            foreach (var item in voltageClassPoints)
            {
                headerInputColumnNumbers.Add(item, new ColNumFormulaUid(totalColumnsCount++));
            }

            headerColumnNumbers.Add(HeaderBal0LogicalParts.OutputSumm, totalColumnsCount++);

            foreach (var item in voltageClassPoints)
            {
                headerOutputColumnNumbers.Add(item, new ColNumFormulaUid(totalColumnsCount++));
            }

            headerColumnNumbers.Add(HeaderBal0LogicalParts.SaldoSumm, totalColumnsCount);
        }

        private TFlxFormat getDefaultFormat;
        private int? headerFormat;
        public int GetHeaderFormat(ExcelFile xls)
        {
            if (headerFormat == null)
            {
                headerFormat = xls.DefaultFormatId;
                getDefaultFormat = xls.GetDefaultFormat;
                getDefaultFormat.Font.Color = Color.Black;
                getDefaultFormat.VAlignment = TVFlxAlignment.center;
                getDefaultFormat.HAlignment = THFlxAlignment.center;
                getDefaultFormat.WrapText = true;
                headerFormat = xls.AddFormat(getDefaultFormat);
            }

            return headerFormat.Value;
        }

        private int? sectionNameFormat;
        public int GetSectionNameFormat(ExcelFile xls)
        {
            if (sectionNameFormat == null)
            {
                sectionNameFormat = xls.DefaultFormatId;
                getDefaultFormat = xls.GetDefaultFormat;
                getDefaultFormat.Font.Style = TFlxFontStyles.Bold;
                getDefaultFormat.Font.Color = Color.Black;
                getDefaultFormat.FillPattern.Pattern = TFlxPatternStyle.Solid;
                getDefaultFormat.FillPattern.FgColor = Color.Silver;
                getDefaultFormat.HAlignment = THFlxAlignment.left;
                sectionNameFormat = xls.AddFormat(getDefaultFormat);
            }

            return sectionNameFormat.Value;
        }

        private int? footnoteFormat;
        public int GetFootnoteFormat(ExcelFile xls)
        {
            if (footnoteFormat == null)
            {
                footnoteFormat = xls.DefaultFormatId;
                getDefaultFormat = xls.GetDefaultFormat;
                getDefaultFormat.HAlignment = THFlxAlignment.left;
                xls.AddFormat(getDefaultFormat);

                getDefaultFormat.HAlignment = THFlxAlignment.right;
                footnoteFormat = xls.AddFormat(getDefaultFormat);
            }

            return footnoteFormat.Value;
        }

        private int? summationFormat;
        public int GetSummationFormat(ExcelFile xls)
        {
            if (summationFormat == null)
            {
                getDefaultFormat = xls.GetDefaultFormat;
                getDefaultFormat.Font.Style = TFlxFontStyles.Bold;
                summationFormat = xls.AddFormat(getDefaultFormat);
            }

            return summationFormat.Value;
        }
    }

    class FormulaIntervalUids
    {
        public FormulaIntervalUids()
        {
            SummInForPs = Guid.NewGuid();
            SummOutForPs = Guid.NewGuid();
            SummSaldoForPs = Guid.NewGuid();
            SaldoByTi = Guid.NewGuid();
            //SummInPssForBalPart = Guid.NewGuid();
            //SummOutPssForBalPart = Guid.NewGuid();
            //SummSaldoPssForBalPart = Guid.NewGuid();
            SummBalPartSaldoForLosses = Guid.NewGuid();
            SummBalPartSaldoForLossesDivider = Guid.NewGuid();
        }
        public Guid SummInForPs { get; set; }
        public Guid SummOutForPs { get; set; }
        public Guid SummSaldoForPs { get; set; }
        public Guid SaldoByTi { get; set; }
        //public Guid SummInPssForBalPart { get; set; }
        //public Guid SummOutPssForBalPart { get; set; }
        //public Guid SummSaldoPssForBalPart { get; set; }
        public Guid SummBalPartSaldoForLosses { get; set; }
        public Guid SummBalPartSaldoForLossesDivider { get; set; }
    }
}
