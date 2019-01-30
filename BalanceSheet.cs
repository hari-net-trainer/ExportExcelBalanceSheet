using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLSample
{
    public class BalanceSheetViewModel
    {
        public string CondonName { get; set; }
        public string CondoAddress { get; set; }
        public string FYPeriod { get; set; }
        public string CondoCity { get; set; }

        public List<CatagoryItems> Liabilities { get; set; }
        public List<CatagoryItems> Assets { get; set; }

        public static BalanceSheetViewModel GetData()
        {
            var rslt = new BalanceSheetViewModel();
            rslt.Liabilities = new List<CatagoryItems>();
            rslt.Assets = new List<CatagoryItems>();

            for (int i = 1; i <= 5; i++)
            {
                rslt.Liabilities.Add(new CatagoryItems { CatagoryName = "XXX", AssetItem = string.Format("Liable - {0}", i), Ammount = 123 * i });
            }
            for (int i = 1; i <= 3; i++)
            {
                rslt.Liabilities.Add(new CatagoryItems { CatagoryName = "YYY", AssetItem = string.Format("Liable - {0}", i), Ammount = 123 * i });
            }

            for (int i = 1; i <= 5; i++)
            {
                rslt.Assets.Add(new CatagoryItems { CatagoryName = "AAA", AssetItem = string.Format("Asset - {0}", i), Ammount = 255 * i });
            }
            for (int i = 1; i <= 3; i++)
            {
                rslt.Assets.Add(new CatagoryItems { CatagoryName = "BBB", AssetItem = string.Format("Asset - {0}", i), Ammount = 255 * i });
            }
            //for (int i = 1; i <= 7; i++)
            //{
            //    rslt.Assets.Add(new CatagoryItems { CatagoryName = "CCC", AssetItem = string.Format("Asset - {0}", i), Ammount = 147 * i });
            //}

            return rslt;
        }

        public static dynamic GetFiscalYearBalanaceSheet()
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
                    oLiab.Items = resultDs.Liabilities.Where(x => x.CatagoryName == cat).Select(y => new { ItemName = y.AssetItem, Value = CurrencyFormate(y.Ammount) });
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
                    oAsst.Items = resultDs.Assets.Where(x => x.CatagoryName == cat).Select(y => new { ItemName = y.AssetItem, Value = CurrencyFormate(y.Ammount) });
                    assetObj.Add(oAsst);
                }
                balObj.AssetItem = new
                {
                    Assets = assetObj,
                    TotalAmount = CurrencyFormate(totalAssetAmt)
                };

                return balObj;
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

    public class CatagoryItems
    {
        public string CatagoryName { get; set; }
        public string AssetItem { get; set; }
        public decimal Ammount { get; set; }
    }

   
}
