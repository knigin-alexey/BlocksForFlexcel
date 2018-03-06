using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Proryv.AskueARM2.Both.VisualCompClasses.Model.Formula;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    public abstract class AbstractBlock<InternalClass> : ISelfRenderable<InternalClass>
    {
        Dictionary<Guid, ExcelFormula> formulaUidDict;
        protected AbstractBlock<InternalClass> Parent;
        protected List<AbstractBlock<InternalClass>> InnerBlocks;
        protected IAffectToNestingLevelBehavior AffectToNestingLevelBehavior { get; set; }

        public AbstractBlock()
        {
            //nestingLevel = 0 может быть только у блока инициализации
            nestingLevel = 0;
            formulaUidDict = new Dictionary<Guid, ExcelFormula>();
            //по умолчанию блок не влияет на уровень вложенности
            AffectToNestingLevelBehavior = new ConcreteNotAffectToNestingLevelBehavior();
        }

        public AbstractBlock(InternalClass internalData) : this()
        {
            this.internalData = internalData;
        }

        public int StartRow { get; set; }
        public int StartCol { get; set; }

        public int UsedRows { get; set; }
        public int UsedCols { get; set; }

        InternalClass internalData;
        public InternalClass InternalData
        {
            get
            {
                if (internalData == null)
                {
                    if (Parent != null)
                    {
                        internalData = Parent.InternalData;
                    }
                }
                return internalData;
            }
            set
            {
                internalData = value;
            }
        }

        private int nestingLevel;
        /// <summary>
        /// Данное свойство нужно для динамического позиционирования элементов в блоке
        /// </summary>
        public int NestingLevel
        {
            get
            {
                return nestingLevel;
            }
        }

        public void AddBlock(AbstractBlock<InternalClass> block)
        {
            if (Object.ReferenceEquals(this, block))
            {
                throw new Exception("Попытка добавления блок в самого себя, зацикливание. Алгоритм для обработки такого случая пока не реализован.");
            }
            if (InnerBlocks == null)
            {
                InnerBlocks = new List<AbstractBlock<InternalClass>>();
            }
            InnerBlocks.Add(block);
            block.SetParent(this);
            //block.SetInternalData(InternalData);
        }

        public void RemoveBlock(AbstractBlock<InternalClass> block)
        {
            InnerBlocks.Remove(block);
            block.SetParent(null);
            block.SetInternalData(default(InternalClass));
            if (AffectToNestingLevelBehavior.IsAffectToNestingLevel())
            {
                block.SetNestingLevel(0);
            }
        }

        public void AddCellToFormula(Guid formulaUid, int col, int row, double value, EnumExcelFormulaOperators oper)
        {
            ExcelFormula formula;
            if (!formulaUidDict.ContainsKey(formulaUid))
            {
                formula = new ExcelFormula();
                formulaUidDict.Add(formulaUid, formula);
            }
            else
            {
                formula = formulaUidDict[formulaUid];
            }

            formula.AddElement(new FormulaElement(col, row, value, oper));
        }

        //public void AddFormula(Guid uid, ExcelFormula formula)
        //{
        //    if (!formulaUidDict.ContainsKey(uid))
        //    {
        //        formulaUidDict.Add(uid, formula);
        //    }
        //    else
        //    {
        //        formulaUidDict[uid].AddElementsRange(formula.Elements);
        //    }
        //}

        public int GetInnerNestingLevel()
        {
            if (InnerBlocks != null && InnerBlocks.Count > 0)
            {
                int maxNesting = 0;

                foreach (var innerNesting in InnerBlocks)
                {
                    if (innerNesting.IsBlockAffectToNestingLevel())
                    {
                        innerNesting.SetNestingLevel(NestingLevel + 1);
                    }
                    int currInnerNesting = innerNesting.GetInnerNestingLevel();
                    if (currInnerNesting > maxNesting)
                    {
                        maxNesting = currInnerNesting;
                    }
                }

                return maxNesting;
            }
            else
            {
                return NestingLevel;
            }
        }

        public int GetMaxParentMaxNestingLevel()
        {
            if (Parent != null)
            {
                return Parent.GetMaxParentMaxNestingLevel();
            }
            else
            {
                return GetInnerNestingLevel();
            }
        }

        public int GetStartRow()
        {
            return StartRow;
        }

        public int GetUsedCols()
        {
            return UsedCols;
        }

        public int GetUsedRows()
        {
            return UsedRows;
        }

        public bool HasInnerBlocks()
        {
            return (InnerBlocks != null && InnerBlocks.Count > 0);
        }

        public bool IsBlockAffectToNestingLevel()
        {
            return AffectToNestingLevelBehavior.IsAffectToNestingLevel();
        }

        public abstract void Render(ExtendedXlsFile xls);

        protected void RenderInnerBlocks(ExtendedXlsFile xls)
        {
            if (InnerBlocks != null && InnerBlocks.Count > 0)
            {
                foreach (var block in InnerBlocks)
                {
                    block.Render(xls);

                    UsedRows = block.GetUsedRows();

                    if (block.GetUsedCols() > UsedCols)
                        UsedCols = block.GetUsedCols();
                }
            }
        }

        public void SetInternalData(InternalClass internalData)
        {
            this.InternalData = internalData;
        }

        public void SetNestingLevel(int nestingLevel)
        {
            this.nestingLevel = nestingLevel;
        }

        public void SetParent(AbstractBlock<InternalClass> parent)
        {
            Parent = parent;
        }

        #region Excel formulas
        protected ExcelFormula GetFormula(Guid uid)
        {
            if (formulaUidDict.ContainsKey(uid))
            {
                return formulaUidDict[uid];
            }
            else
            {
                return new ExcelFormula();
            }
        }

        protected ExcelFormula GetAllInnerFormulas(Guid uid)
        {
            ExcelFormula innerFormulasUnited = new ExcelFormula();
            if (InnerBlocks != null && InnerBlocks.Count > 0)
            {
                foreach (var innerBlock in InnerBlocks)
                {
                    ExcelFormula innerFormula = innerBlock.GetAllInnerFormulas(uid);
                    innerFormulasUnited.AddElementsRange(innerFormula.Elements);
                }
            }
            else
            {
                innerFormulasUnited = GetFormula(uid);
            }

            return innerFormulasUnited;
        }

        protected ExcelFormula GetTopParentFirstEntryFormulas(Guid uid)
        {
            ExcelFormula innerFormulasUnited = new ExcelFormula();
            if (Parent != null)
            {
                return Parent.GetTopParentFirstEntryFormulas(uid);
            }
            else
            {
                innerFormulasUnited.AddElementsRange(GetInnerFirstEntryFormulas(uid).Elements);
                return innerFormulasUnited;
            }
        }

        //TODO: Похоже на проблему, дублирование функционала
        protected ExcelFormula GetInnerFirstEntryFormulas(Guid uid)
        {
            ExcelFormula innerFormulasUnited = new ExcelFormula();
            if (InnerBlocks != null && InnerBlocks.Count > 0)
            {
                foreach (var innerBlock in InnerBlocks)
                {
                    ExcelFormula innerFormula = innerBlock.GetFormula(uid);
                    if (innerFormula.Elements.Count == 0)
                    {
                        innerFormula = innerBlock.GetInnerFirstEntryFormulas(uid);
                    }

                    innerFormulasUnited.AddElementsRange(innerFormula.Elements);
                }
            }

            return innerFormulasUnited;
        }

        //TODO: Похоже на проблему, дублирование функционала
        protected ExcelFormula GetFirstLevelFormulas(Guid uid)
        {
            ExcelFormula innerFormulasUnited = new ExcelFormula();
            if (InnerBlocks != null && InnerBlocks.Count > 0)
            {
                foreach (var innerBlock in InnerBlocks)
                {
                    ExcelFormula innerFormula = innerBlock.GetFormula(uid);
                    innerFormulasUnited.AddElementsRange(innerFormula.Elements);
                }
            }

            return innerFormulasUnited;
        }

        protected void RemoveFormulaById(Guid ID)
        {
            if (formulaUidDict.ContainsKey(ID))
            {
                formulaUidDict.Remove(ID);
            }
            else
            {
                throw new IndexOutOfRangeException("Попытка удалить формулу из ListFormulsByID по несуществующему ключу.");
            }
        }
        #endregion Excel formulas
    }
}
