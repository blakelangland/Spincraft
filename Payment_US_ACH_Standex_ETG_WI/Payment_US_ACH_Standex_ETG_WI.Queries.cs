using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using Epicor.Data;
using Epicor.Utilities;
using Erp.Tables;

namespace Erp.Internal.EI
{
    public partial class Payment_US_ACH_Standex_ETG_WI
    {
        #region APTran queries
        private static Func<ErpContext, string, int, IEnumerable<APTran>> selectAPTranQuery;
        private IEnumerable<APTran> SelectAPTran(string company, int headNum, int maxCount)
        {
            if (selectAPTranQuery == null)
            {
                Expression<Func<ErpContext, string, int, IEnumerable<APTran>>> Expression =
                  (ctx, company_ex, headNum_ex) =>
                    (from row in ctx.APTran
                     where row.Company == company_ex &&
                      row.TranType == "Pay" &&
                      row.HeadNum == headNum_ex
                     orderby row.InvoiceNum
                     select row);
                selectAPTranQuery = DBExpressionCompiler.Compile(Expression);
            }
            return selectAPTranQuery(this.Db, company, headNum).Take(maxCount);
        }


        private class APTranPartialRow
        {
            public string InvoiceNum { get; set; }
        }

        private static Func<ErpContext, string, int, APTranPartialRow> findFirstAPTranQuery;
        private APTranPartialRow FindFirstAPTranQuery(string company, int headNum)
        {
            if (findFirstAPTranQuery == null)
            {
                Expression<Func<ErpContext, string, int, APTranPartialRow>> Expression =
                  (ctx, company_ex, headNum_ex) =>
                    (from row in ctx.APTran
                     where row.Company == company_ex &&
                            row.HeadNum == headNum_ex
                     select new APTranPartialRow
                     {
                         InvoiceNum = row.InvoiceNum
                     }).FirstOrDefault();
                findFirstAPTranQuery = DBExpressionCompiler.Compile(Expression);
            }
            return findFirstAPTranQuery(this.Db, company, headNum); ;
        }


        #endregion

        #region VendBank queries
        private class VendBankResult
        {
            public string VendorID { get; set; }
            public int VendorNum { get; set; }
            public string VendorBankID { get; set; }            
            public int BCountryNum { get; set; }
            public string ISOCode { get; set; }
        }

        private static Func<ErpContext, string, IEnumerable<VendBankResult>> vendBankQuery;
        private IEnumerable<VendBankResult> VendBankQuery(string company)
        {
            if (vendBankQuery == null)
            {
                Expression<Func<ErpContext, string, IEnumerable<VendBankResult>>> Expression =
                  (ctx, company_ex) =>
                    (from row in ctx.VendBank
                     join vendor_Row in ctx.Vendor 
                     on new { row.Company, row.VendorNum } 
                     equals new { vendor_Row.Company, vendor_Row.VendorNum }
                     from country_Row in ctx.Country.Where(r => r.Company == row.Company &&
                                                                r.CountryNum == row.BCountryNum).DefaultIfEmpty()
                     where row.Company == company_ex
                     select new VendBankResult()
                     {
                         VendorNum = row.VendorNum,
                         VendorID = vendor_Row.VendorID,
                         VendorBankID = row.BankID,
                         BCountryNum = row.BCountryNum,
                         ISOCode = country_Row != null  ? country_Row.ISOCode : string.Empty
                     }).Distinct();
                vendBankQuery = DBExpressionCompiler.Compile(Expression);
            }
            return vendBankQuery(this.Db, company);
        }
        #endregion
    }
}
