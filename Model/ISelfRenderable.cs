using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.Classes;
using Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    public interface ISelfRenderable<Internal>
    {
        void SetParent(AbstractBlock<Internal> parent);
        void SetNestingLevel(int nestLev);
        void SetInternalData(Internal internalData);
        Internal InternalData { get; }
        void Render(ExtendedXlsFile xls);
        int GetUsedRows();
        int GetUsedCols();
        int GetStartRow();
        int GetInnerNestingLevel();
        int GetMaxParentMaxNestingLevel();
        //ExcelFormula GetFormula(Guid uid);
        //ExcelFormula GetInnerFirstEntryFormulas(Guid uid);
        //void AddCellToFormula(Guid id, int col, int row, double value, EnumExcelFormulaOperators oper);
        bool IsBlockAffectToNestingLevel();
        void AddBlock(AbstractBlock<Internal> block);
        bool HasInnerBlocks();
        //void AddFormula(Guid uid, ExcelFormula formula);
        //ExcelFormula GetAllInnerFormulas(Guid uid);
    }
}
