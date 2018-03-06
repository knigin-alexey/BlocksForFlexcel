using FlexCel.Core;
using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.Classes;
using Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    /// <summary>
    /// Пока задумка такая: блок - это некий абстрактный блок в экселе/html.
    /// Максимальный блок - весь документ, минимальный блок - одна строка
    /// Идея для будущего: блоки - это части ExtendedXlsFile, а внутри него 
    /// происходит расчет StartRow, UsedRow и прочее. И автоматическая инкрементация
    /// </summary>
    public abstract class AbstractDataBlock<DataClass, InternalClass> : AbstractBlock<InternalClass>
    {
        protected Color IsOVReplacedColor = Color.Red;
        protected Color ZeroCrossDetected = Color.Red;
        protected Color IsManualStatusChangeColor = Color.Yellow;
        protected Color NoDataColor = Color.LightPink;
        protected Color NoDrumsColor = Color.DarkOrange;

        private DataClass data;
        public DataClass Data
        {
            get
            {
                return data;
            }
        }

        public AbstractDataBlock(DataClass data) : base()
        {
            this.data = data;
            
        }

        public AbstractDataBlock(DataClass data, InternalClass internalData) : base(internalData)
        {
            this.data = data;
        }

        #region Border
        public void SetBorderAllCellsInBlock(ExtendedXlsFile xls, Color borderColor, TFlxBorderStyle linestyle)
        {
            xls.SetBorderRangeCells(StartRow, StartCol, UsedRows, UsedCols, borderColor, linestyle);
        }

        public void SetBlockBorderBySide(ExtendedXlsFile xls, Color borderColor, TFlxBorderStyle linestyle, SideToBorder side)
        {
            xls.SetCellRangeBorderBySide(StartRow, StartCol, UsedRows, UsedCols, borderColor, linestyle, side);
        }
        #endregion Border
    }
}
