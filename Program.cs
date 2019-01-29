using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLSample
{
    class Program
    {
        static void Main(string[] args)
        {
            Report report = new Report();

            report.CreateExcelDoc(@"E:\d drive\Hari Krishna\Report.xlsx");

            Console.WriteLine("Excel file has created!");
        }

        public static void GetFiscalYearBalanaceSheet()
        {
            string fiscalYear = string.Empty;
            decimal totalAssetAmt = 0, totalLiabAmt = 0;
            dynamic balObj = new ExpandoObject();
            try
            {
                balObj.CondoName = "Sushan Condo";
                balObj.CondoAdd = "Some Address";
                balObj.CondoCity = "City Name";
                balObj.FYPeriod = "FY2018-19";

                var resultDs = BalanceSheetViewModel.GetData();

                //distinct values of Catagory
                var lstLiabCat = resultDs.Liabilities.Select(x => x.CatagoryName).Distinct();
                totalLiabAmt = resultDs.Liabilities.Sum(x => x.Ammount);
                List<dynamic> liabObj = new List<dynamic>();
                foreach (var cat in lstLiabCat)
                {
                    dynamic oLiab = new ExpandoObject();
                    oLiab.CatagoryItem = cat;
                    oLiab.TotalValue = CurrencyFormate(resultDs.Liabilities.Where(x => x.CatagoryName == cat).Sum(y => y.Ammount));
                    oLiab.Items = resultDs.Liabilities.Where(x => x.CatagoryName == cat).Select(y => new { Item = y.AssetItem, Value = CurrencyFormate(y.Ammount) });
                    liabObj.Add(oLiab);
                }
                balObj.LiabilityItem = new
                {
                    Liabilities = liabObj,
                    TotalAmount = CurrencyFormate(totalLiabAmt)
                };

                var lstAssetCat = resultDs.Assets.Select(x => x.CatagoryName).Distinct();
                totalAssetAmt = resultDs.Assets.Sum(x => x.Ammount);
                List<dynamic> assetObj = new List<dynamic>();
                foreach (var cat in lstAssetCat)
                {
                    dynamic oAsst = new ExpandoObject();
                    oAsst.CatagoryItem = cat;
                    oAsst.TotalValue = CurrencyFormate(resultDs.Assets.Where(x => x.CatagoryName == cat).Sum(y => y.Ammount));
                    oAsst.Items = resultDs.Assets.Where(x => x.CatagoryName == cat).Select(y => new { Item = y.AssetItem, Value = CurrencyFormate(y.Ammount) });
                    assetObj.Add(oAsst);
                }
                balObj.AssetItem = new
                {
                    Assets = assetObj,
                    TotalAmount = CurrencyFormate(totalAssetAmt)
                };

                }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static string CurrencyFormate(decimal amount)
        {
            var retVal = amount < 0 ? string.Format("({0:N2})", amount * -1) : string.Format("{0:N2}", amount);
            return retVal;
        }

    }
}
