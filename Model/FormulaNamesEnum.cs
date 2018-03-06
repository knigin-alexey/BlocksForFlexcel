using Proryv.AskueARM2.Server.DBAccess.Internal.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Proryv.AskueARM2.Both.VisualCompClasses.Model
{
    public enum FormulaNamesEnum
    {
        [ExcelFormulaMapper("bd5868ce-9014-4b66-a08e-3fe589ad6d0d")]
        ConsolidatedActSaldo220,
        [ExcelFormulaMapper("66284469-61ca-426f-bf0a-c9ee28e271ff")]
        ConsolidatedActSaldo330,
        [ExcelFormulaMapper("4f143d0a-2d47-4c03-8c52-107fd34bed18")]
        ConsolidatedActSaldoSummary,
        [ExcelFormulaMapper("C4CC8E2F-541D-4BB9-AE1F-46F71CBE4414")]
        Balance220330RelativeLosses,
        [ExcelFormulaMapper("B4944D18-53DE-4ABC-A5D2-D0D349BF70FE")]
        Balance220330InputSummary,
        [ExcelFormulaMapper("AEC127B9-94FB-4F75-83F9-5DC12FEA2245")]
        Balance220330OutputSummary,
        [ExcelFormulaMapper("76D721D9-4A44-4FEB-84B7-77444D93B37A")]
        Balance220330SaldoSummary,
        [ExcelFormulaMapper("D2E26C76-3A26-4D59-8899-9ADA81201E75")]
        Balance220330AllSaldoSummary,
        [ExcelFormulaMapper("BC20381F-277F-41D0-ABE3-2CC02A5C336A")]
        Balance220330AllOutputSummary,
        [ExcelFormulaMapper("6C9EAFA5-8C79-40BA-8A56-B92C779C8BFD")]
        BalanceEnesSaldoSummary,
        [ExcelFormulaMapper("051A5BAD-C6A3-4DD8-9414-EFAEEF48E395")]
        BalanceEnesInputBalPartSummary,
        [ExcelFormulaMapper("279CC54F-E3D0-4D37-8D5B-CCBB51717320")]
        BalanceEnesOutputBalPartSummary,
        [ExcelFormulaMapper("35680EA7-85F2-4598-BB47-8865B809E01D")]
        BalanceEnesInputTotalSummary,
        [ExcelFormulaMapper("64B75856-15D5-4CBC-AD18-D29F68217D79")]
        BalanceEnesOutputTotalSummary,
        [ExcelFormulaMapper("1C6D944E-0C33-48C2-9EBE-CBF962CFEBF2")]
        BalanceEnesSaldoTotalSummary,
        //[ExcelFormulaMapper("6B3DAB8E-7BF5-43DE-A76F-EB451E0950D5")]
        //BalancePsTransSummary,
        //[ExcelFormulaMapper("E5BF20C8-A3CF-48C5-9761-C2753FCBC8EB")]
        //BalancePsReactSummary,
        [ExcelFormulaMapper("C8E21CAC-52E4-4000-86A3-CB9591B5C6D8")]
        BalancePsOtpuskSShin2ByVoltage,
        [ExcelFormulaMapper("6A512EE9-F69C-4247-AB4E-33A00F7BD899")]
        BalancePsValueByDrumSum,
        [ExcelFormulaMapper("895B7438-6691-4E4F-9294-E4E68604CF06")]
        IASaldoDrum,
        [ExcelFormulaMapper("4F8CABCF-1265-404F-80BB-E56DA777A9F5")]
        IASaldo330Drum,
        [ExcelFormulaMapper("97D0D019-43A4-4ED5-AB3E-5179D617554D")]
        IASaldo330Border,
        [ExcelFormulaMapper("214EC82C-28F4-4FE5-8B3A-F08922FA812F")]
        IASaldo220Drum,
        [ExcelFormulaMapper("BBCEC6BA-4FF0-4D2F-B1D8-A398B51EDDF4")]
        IASaldo220Border,
        [ExcelFormulaMapper("74D67B80-48BF-4F3B-A2EB-9196C7B90617")]
        IAEnesDrumInput,
        [ExcelFormulaMapper("7585F531-29D0-4E9A-95C8-059B51D3072B")]
        IAEnesBorderInput,
        [ExcelFormulaMapper("8AC7D891-ECA7-43EF-841C-0A9F15964421")]
        IAEnesDrumOutput,
        [ExcelFormulaMapper("459B406B-2D07-4F18-AF99-58364508E400")]
        IAEnesBorderOutput,
        [ExcelFormulaMapper("8066C0FC-537D-4FB0-B2AD-0BD78A4E0DCF")]
        IAEnesDrumSaldo,
        [ExcelFormulaMapper("B23BB1A5-1709-48DD-A0A1-1C7796EC0EC2")]
        IAEnesBorderSaldo,
        [ExcelFormulaMapper("83E83102-3E20-400F-ADD0-173F367B0771")]
        IAFullConsAct220Saldo,
        [ExcelFormulaMapper("AFFAD3CD-8EE6-4CF9-B590-57C5F41F9315")]
        IAFullConsAct330Saldo,
        [ExcelFormulaMapper("2B74D878-CF20-4779-B4A5-31267E1A8027")]
        IaIntervalOutputDrum220,
        [ExcelFormulaMapper("B1C336AA-66C9-4933-A551-376860B9D285")]
        IaIntervalOutputDrum330,
        [ExcelFormulaMapper("E749800B-D8D6-43B9-A753-C097A08782F4")]
        IaIntervalInputDrum220,
        [ExcelFormulaMapper("8127280D-8AE6-47AC-B2E4-C13BEF6EE4CE")]
        IaIntervalInputDrum330
    }
}
