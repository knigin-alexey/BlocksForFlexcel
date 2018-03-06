using Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Data.Full;
using Proryv.AskueARM2.Both.VisualCompClasses.Model;
using Proryv.AskueARM2.Server.DBAccess.Internal;
using Proryv.AskueARM2.Server.DBAccess.Internal.TClasses;
using Proryv.AskueARM2.Server.DBAccess.Public.Calculation.Balances;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Concrete.BalansHierLev0.Full
{
    public class BalanceHierLev0RenderFullBehavior
    {
        public ExtendedXlsFile TestExecuteBalansHierLev0_Valtage_220_330_Full(BalansHierLev0Result HierLev0,
            TExportExcelAdapterType ExportType, IVisualDataRequestObjectsNames getNameInterface)
        {
            Dictionary<ID_TypeHierarchy, string> dictionaryOfNames = HierLev0.DictionaryOfNames;
            InitData data = new InitData(HierLev0.VoltageClassPoints);
            InternalData internalData = new InternalData(ExportType);
            InitBlock initBlock = new InitBlock(data, internalData);
            Classes.TitleInfo titleData =
                new Classes.TitleInfo(getNameInterface.GetBalanceNameForHierLev0(HierLev0.BalanceId), HierLev0.DTStart, HierLev0.DTEnd);
            TitleBlock titleBlock = new TitleBlock(titleData);
            initBlock.AddBlock(titleBlock);
            HeaderFooterData headFootData = new HeaderFooterData(HierLev0.VoltageClass, HierLev0.HighLimit);
            HeaderFooterBlock headerFooter = new HeaderFooterBlock(headFootData, internalData);
            initBlock.AddBlock(headerFooter);

            foreach (TIntegral_HierLev0_Values balanceSection in HierLev0.Result_Values)
            {
                BalancePartData balData = new BalancePartData(getNameInterface.GetBalanceSectionName(balanceSection.HierLev0Group_Name),
                    HierLev0.BalPartList.Where(x => x.IsUseInGeneralBalance).Select(x => x.Name).Contains(balanceSection.HierLev0Group_Name),
                    HierLev0.BalPartList.Where(x => x.IsRsk).Select(x => x.Name).Contains(balanceSection.HierLev0Group_Name));

                BalancePartBlock balPartBlock = new BalancePartBlock(balData, internalData);
                foreach (KeyValuePair<ID_IsOurSide, TIntegral_PS_ValuesForHierLev0> psBalSect in balanceSection.HierLev0DetailGroupResult)
                {
                    string psName;
                    TIntegral_PS_ValuesForHierLev0 psBalSectData = psBalSect.Value;
                    ID_IsOurSide side = psBalSect.Key;
                    ID_TypeHierarchy key = new ID_TypeHierarchy(enumTypeHierarchy.Dict_PS, -1);
                    key.TypeHierarchy = side.IsOurSide ? enumTypeHierarchy.Dict_PS : enumTypeHierarchy.Dict_Contr_PS;
                    key.ID = side.ID;

                    if (!dictionaryOfNames.TryGetValue(key, out psName))
                    {
                        psName = getNameInterface.GetPSName(side.ID, !side.IsOurSide);
                    }
                    PsBlock psBlock = new PsBlock(psName, internalData);
                    foreach (TI_Integral_ValuesForHierLev0 tiPsBalSect in psBalSectData.TI_List)
                    {
                        string tIName = string.Empty;
                        key.TypeHierarchy = tiPsBalSect.TypeHierarchy;
                        key.ID = tiPsBalSect.ID;
                        if (!dictionaryOfNames.TryGetValue(key, out tIName))
                        {
                            switch (tiPsBalSect.TypeHierarchy)
                            {
                                case enumTypeHierarchy.Dict_PS:
                                    tIName = getNameInterface.GetPSName(tiPsBalSect.ID, !side.IsOurSide);
                                    break;

                                case enumTypeHierarchy.Info_TI:
                                    tIName = getNameInterface.GetTIName(tiPsBalSect.ID, false);
                                    break;

                                case enumTypeHierarchy.Info_ContrTI:
                                    tIName = getNameInterface.GetTIName(tiPsBalSect.ID, true);
                                    break;

                                case enumTypeHierarchy.Info_TP:
                                    tIName = getNameInterface.GetTPName(tiPsBalSect.ID);
                                    break;
                            }
                        }

                        List<TVALUES_DB> inputValues;
                        List<TVALUES_DB> outputValues;
                        Data.Full.TiData tiData = new Data.Full.TiData(tIName);
                        if (tiPsBalSect.Val_List.TryGetValue(1, out inputValues))
                        {
                            tiData.InputByVoltages.Add(tiPsBalSect.VoltageClassPoint, inputValues[0]);
                        }
                        if (tiPsBalSect.Val_List.TryGetValue(2, out outputValues))
                        {
                            tiData.OutputByVoltages.Add(tiPsBalSect.VoltageClassPoint, outputValues[0]);
                        }

                        TiBlock tiBlock = new TiBlock(tiData, internalData);
                        psBlock.AddBlock(tiBlock);
                    }
                    balPartBlock.AddBlock(psBlock);
                }
                headerFooter.AddBlock(balPartBlock);
            }

            ExtendedXlsFile xls = new ExtendedXlsFile(ExportType);
            initBlock.Render(xls);

            return xls;
        }
    }
}
