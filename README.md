# BlocksForFlexcel

Все блоки унаследованы от классов AbstractBodyBlock<T, H> и AbstractBodyFooterBlock<T, H>, которые унаследованы от AbstractBlock<T, H> : ISelfRenderable<H>

Дженерик H - это тот самый класс InternalData, который общий для всей ветки.
Он был добавлен для того, чтобы у разного вида отчетов были разные хранилища всяких кастомных данных (или, например, костылей). В моем случае, в InternalData хранятся Dictionary<Enum, int>, где енамы - это наименования столбцов, а int - это номер колонки в экселе. Причем, словарей таких аж 3 - HeaderColumnNumbers, HeaderInputColumnNumbers, HeaderOutputColumnNumbers. 
HeaderColumnNumbers сделан из-за того, что ТИ может оказаться в колонке №3 или №4 из-за вложенности (в арм 3.0 ПС может быть вложена не только в раздел, но еще и в субъект РФ), а всем остальным колонкам придется сменить номер.
HeaderInputColumnNumbers, HeaderOutputColumnNumbers - динамические, зависят от количества типов вольтажей в коллекции ТИ. 
Инициализация этих словарей происходит в методе Initialize(), который вызывается при отрисовке самого верхнего блока initBlock. Почему? Потому что все блоки уже добавлены и можно определить позиции всех элементов.
На месте InternalData можно использовать любой, на то нам и дженерик.

В AbstractBlock реализован следующий функционал:
Цвета переворотов, обходников и прочего для экселя (наверное, не очень удачно и можно это запихать в ExtendedXlsFile).
Словарь Dictionary<Guid, ExcelFormula> formulaUidDict (про это мы поговорим отдельно)
Еще один дженерик <T> data - вот этот класс может быть свой у каждого блока

Коллекция List<ISelfRenderable<H>> InnerBlocks
Уровень вложенности int NestingLevel
Поведение IAffectToNestingLevelBehavior AffectToNestingLevelBehavior (влияет или нет на уровень вложенности, допустим header в нашем случае не влияет, а раздел баланса уже начинает влиять, по умолчанию - не влияет)
Координаты блока: StartRow, StartCol, UsedRows (по этому свойству происходит основная итерация строк, типа, везде пишу UsedRows++ и пишу в новой строке), UsedCols

Основной метод
public abstract void Render(ExtendedXlsFile xls);

Два конструктора.
Один принимает только данные, а потом InternalData берет из родительского объекта, что очень не очевидно и косячно.
Второй принимает и данные и InternalData.

AddBlock(ISelfRenderable<H> block)
Добавляет block в InnerBlocks и задает block'у родителя (через метод SetParent)

RemoveBlock(ISelfRenderable<H> block)
Удаляет блок из коллекции 

int GetInnerNestingLevel()
Возвращает максимальный NestingLevel, опрашивая всех детей

int GetMaxParentMaxNestingLevel()
Возвращает максимальный NestingLevel, опрашивая самого главного родителя

void RenderInnerBlocks(ExtendedXlsFile xls)
в цикле обходит InnerBlocks и вызывает у каждого абстрактный Render, потом получает из него GetUsedRows и записывает его в свои UsedRows

геттеры, сеттеры
GetUsedRows(), GetUsedCols(), SetParent(ISelfRenderable<H> parent),  SetNestingLevel(int nestingLevel)

Методы
SetBorderAllCellsInBlock(ExtendedXlsFile xls, Color borderColor, TFlxBorderStyle linestyle) - выставляем границу во всех ячейках в блоке
SetBlockBorderBySide(ExtendedXlsFile xls, Color borderColor, TFlxBorderStyle linestyle, SideToBorder side) - выставляем границу по определенной стороне блока (или по всем, делаем рамку)

Методы и поля, относящиеся к формированию эксель-формул
Dictionary<Guid, ExcelFormula> formulaUidDict;

ExcelFormula GetFormula(Guid uid) 
если есть uid в formulaUidDict, то возвращаем value, иначе new ExcelFormula()

void RemoveFormulaById(Guid ID)
удаляет формулу из словаря

void AddCellToFormula(Guid uid, int col, int row, double value, EnumExcelFormulaOperators oper)
Говорим механизму, что для uid'а формулы такого-то добавить FormulaElement с координатами col, row, оператором oper и значением value. А если такой формулы нет, то её создать и в неё добавить.

GetFirstLevFormulas(Guid uid)
в дочерних блоках первого уровня вложенности опрашиваем formulaUidDict и делаем из него объединенный словарь в вызывающем блоке-родителе
Как используется:
В блоках ТИ мы в колонку вольтажа добавляем значение и добавляем полученные координаты (колонку, строку) и это значение в формулу HeaderBal0LogicalParts.InputSumm
<pre><code class="java">
AddCellToFormula(HeaderBal0LogicalParts.InputSumm.GetFormulaUid(),
  InternalData.HeaderInputColumnNumbers[inputByVolt.Key].ColumnNumber, UsedRows, *inputValue*, EnumExcelFormulaOperators.Plus);
</code></pre>

HeaderBal0LogicalParts - это enum наименования столбцов и в нем кастомными атрибутами заданы уиды формул 
<pre><code class="java">
[ExcelFormulaMapper("89896375-0722-422b-8d19-fe87af0419eb")]
InputSumm,
</code></pre>

То есть для данной ТИ в формулу HeaderBal0LogicalParts.InputSumm добавлена ячейка с колонкой "Прием кВт.ч." и строкой данной ТИ.

После этого мы пишем:
<pre><code class="java">
AddCellToFormula(FormulaNamesEnum.Balance220330InputSummary.GetFormulaUid(),
    InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm], UsedRows,
    inputFormula.DoubleRepresentation(), EnumExcelFormulaOperators.Plus);
</code></pre>

И таким образом мы помещаем в формулу FormulaNamesEnum.Balance220330InputSummary координаты ячейки и в значение мы записываем inputFormula.DoubleRepresentation().
После этого, уже в блоке ПС при отображении вызывается:
<pre><code class="java">
xls.SetFormulaStyled(block.GetUsedRows(),
  block.InternalData.HeaderColumnNumbers[HeaderBal0LogicalParts.InputSumm],
  block.GetFirstLevFormulas(FormulaNamesEnum.Balance220330InputSummary.GetFormulaUid()),
  style);
</code></pre>

Блок ПС формирует единую формулу из всех дочерних блоков ТИ по идентификатору FormulaNamesEnum.Balance220330InputSummary и передает получившийся объект в метод xls.SetFormulaStyled(row, col, excelFormula)
Этот метод, в зависимости от формы отображения документа (xls или html), вызывает у формулы методы StringRepresentation или DoubleRepresentation.

Вот таким образом решается проблема с формулами и аккумуляторами и дублирующимся кодом.

h3. Ну и, собственно, сам механизм отрисовки.

AbstractBodyBlock<T, H> 
Конструкторы
<pre><code class="java">
protected abstract void RenderBody(ExtendedXlsFile xls);

public override void Render(ExtendedXlsFile xls)
{
    RenderBody(xls);
    RenderInnerBlocks(xls);
}
</code></pre>

AbstractBodyFooterBlock<T, H>

Конструкторы
<pre><code class="java">
protected abstract void RenderBody(ExtendedXlsFile xls);
protected abstract void RenderFooter(ExtendedXlsFile xls);

public override void Render(ExtendedXlsFile xls)
{
    RenderBody(xls);
    RenderInnerBlocks(xls);
    RenderFooter(xls);
}
</code></pre>

Блок ТИ унаследован от AbstractBodyBlock, а блоки ПС и разделов баланса от AbstractBodyFooterBlock, потому что им нужно считать и отображать итоги.

Для детального изучения предлагаю запускать в дебаге метод ExcelReportAdapter.TestExecuteBalansHierLev0_Valtage_220_330(...), пока он так называется.
