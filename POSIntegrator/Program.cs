using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using WebApiContrib.Formatting;
using System.IO;
using System.Net;
using System.Web.Script.Serialization;
using System.Diagnostics;
using System.Globalization;

using Excel = Microsoft.Office.Interop.Excel;

namespace POSIntegrator
{
    // =======
    // Program
    // =======
    class Program
    {
        // =============
        // Data Contexts
        // =============
        private static POS13db.POS13dbDataContext posData = new POS13db.POS13dbDataContext();

        // ===================
        // Fill Leading Zeroes
        // ===================
        public static String FillLeadingZeroes(Int32 number, Int32 length)
        {
            var result = number.ToString();
            var pad = length - result.Length;
            while (pad > 0)
            {
                result = '0' + result;
                pad--;
            }

            return result;
        }

        // ============
        // Sync StockIn
        // ============
        public static void SyncStockIn(string xlsPath)
        {
            try
            {
                Console.WriteLine("Scanning XLS Files...");
                Boolean hasFiles = false;

                List<string> files = new List<string>(Directory.EnumerateFiles(xlsPath));
                foreach (var file in files)
                {
                    hasFiles = true;
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    Excel.Range range = xlWorkSheet.UsedRange;

                    Int32 stockInId = 0;

                    Int32 row = range.Rows.Count;
                    Int32 column = range.Columns.Count;

                    Console.WriteLine("Saving " + (range.Cells[2, 1] as Excel.Range).Value2 + "...");
                    Console.WriteLine();

                    for (Int32 rowCount = 1; rowCount <= row; rowCount++)
                    {
                        if (rowCount > 1)
                        {
                            if (rowCount == 2)
                            {
                                String documentReference = (range.Cells[rowCount, 1] as Excel.Range).Value2;
                                String barCode = (range.Cells[rowCount, 2] as Excel.Range).Value2;
                                String item = (range.Cells[rowCount, 3] as Excel.Range).Value2;
                                Decimal quantity = Convert.ToDecimal((range.Cells[rowCount, 4] as Excel.Range).Value2);
                                Decimal cost = Convert.ToDecimal((range.Cells[rowCount, 5] as Excel.Range).Value2);
                                Decimal amount = Convert.ToDecimal((range.Cells[rowCount, 6] as Excel.Range).Value2);

                                var stockIn = from d in posData.TrnStockIns
                                              where d.Remarks.Equals(documentReference)
                                              select d;

                                if (!stockIn.Any())
                                {
                                    var defaultPeriod = from d in posData.MstPeriods
                                                        select d;

                                    var defaultSettings = from d in posData.SysSettings
                                                          select d;

                                    var defaultStockInNumber = defaultPeriod.FirstOrDefault().Period + "-000001";

                                    var lastStockIn = from d in posData.TrnStockIns.OrderByDescending(d => d.Id)
                                                      select d;

                                    if (lastStockIn.Any())
                                    {
                                        var lastStockInNumber = lastStockIn.FirstOrDefault().StockInNumber.ToString();
                                        int secondIndex = lastStockInNumber.IndexOf('-', lastStockInNumber.IndexOf('-'));
                                        var stockInNumberSplitStringValue = lastStockInNumber.Substring(secondIndex + 1);
                                        var stockInNumber = Convert.ToInt32(stockInNumberSplitStringValue) + 000001;
                                        defaultStockInNumber = defaultPeriod.FirstOrDefault().Period + "-" + FillLeadingZeroes(stockInNumber, 6);
                                    }

                                    POS13db.TrnStockIn newStockIn = new POS13db.TrnStockIn
                                    {
                                        PeriodId = defaultPeriod.FirstOrDefault().Id,
                                        StockInDate = DateTime.Today,
                                        StockInNumber = defaultStockInNumber,
                                        SupplierId = defaultSettings.FirstOrDefault().PostSupplierId,
                                        Remarks = documentReference,
                                        IsReturn = false,
                                        CollectionId = null,
                                        PurchaseOrderId = null,
                                        PreparedBy = defaultSettings.FirstOrDefault().PostUserId,
                                        CheckedBy = defaultSettings.FirstOrDefault().PostUserId,
                                        ApprovedBy = defaultSettings.FirstOrDefault().PostUserId,
                                        IsLocked = true,
                                        EntryUserId = defaultSettings.FirstOrDefault().PostUserId,
                                        EntryDateTime = DateTime.Now,
                                        UpdateUserId = defaultSettings.FirstOrDefault().PostUserId,
                                        UpdateDateTime = DateTime.Now,
                                        SalesId = null
                                    };

                                    posData.TrnStockIns.InsertOnSubmit(newStockIn);
                                    posData.SubmitChanges();

                                    stockInId = newStockIn.Id;

                                    var items = from d in posData.MstItems
                                                where d.BarCode.Equals(barCode)
                                                select d;

                                    if (items.Any())
                                    {
                                        POS13db.TrnStockInLine newStockInLine = new POS13db.TrnStockInLine
                                        {
                                            StockInId = stockInId,
                                            ItemId = items.FirstOrDefault().Id,
                                            UnitId = items.FirstOrDefault().UnitId,
                                            Quantity = Convert.ToInt32(quantity),
                                            Cost = Convert.ToDecimal(cost),
                                            Amount = Convert.ToDecimal(amount),
                                            ExpiryDate = items.FirstOrDefault().ExpiryDate,
                                            LotNumber = items.FirstOrDefault().LotNumber,
                                            AssetAccountId = items.FirstOrDefault().AssetAccountId,
                                            Price = items.FirstOrDefault().Price
                                        };

                                        posData.TrnStockInLines.InsertOnSubmit(newStockInLine);

                                        var currentItem = from d in posData.MstItems
                                                          where d.Id == newStockInLine.ItemId
                                                          select d;

                                        if (currentItem.Any())
                                        {
                                            Decimal currentOnHandQuantity = currentItem.FirstOrDefault().OnhandQuantity;
                                            Decimal totalQuantity = currentOnHandQuantity + Convert.ToDecimal(quantity);
                                            var updateItem = currentItem.FirstOrDefault();
                                            updateItem.OnhandQuantity = totalQuantity;
                                        }

                                        posData.SubmitChanges();
                                        Console.WriteLine(item + " was successfuly saved!");
                                    }
                                }
                            }
                            else
                            {
                                if (rowCount > 2)
                                {
                                    String documentReference = (range.Cells[rowCount, 1] as Excel.Range).Value2;
                                    String barCode = (range.Cells[rowCount, 2] as Excel.Range).Value2;
                                    String item = (range.Cells[rowCount, 3] as Excel.Range).Value2;
                                    Decimal quantity = Convert.ToDecimal((range.Cells[rowCount, 4] as Excel.Range).Value2);
                                    Decimal cost = Convert.ToDecimal((range.Cells[rowCount, 5] as Excel.Range).Value2);
                                    Decimal amount = Convert.ToDecimal((range.Cells[rowCount, 6] as Excel.Range).Value2);

                                    if (stockInId > 0)
                                    {
                                        var items = from d in posData.MstItems
                                                    where d.BarCode.Equals(barCode)
                                                    select d;

                                        if (items.Any())
                                        {
                                            POS13db.TrnStockInLine newStockInLine = new POS13db.TrnStockInLine
                                            {
                                                StockInId = stockInId,
                                                ItemId = items.FirstOrDefault().Id,
                                                UnitId = items.FirstOrDefault().UnitId,
                                                Quantity = Convert.ToInt32(quantity),
                                                Cost = Convert.ToDecimal(cost),
                                                Amount = Convert.ToDecimal(amount),
                                                ExpiryDate = items.FirstOrDefault().ExpiryDate,
                                                LotNumber = items.FirstOrDefault().LotNumber,
                                                AssetAccountId = items.FirstOrDefault().AssetAccountId,
                                                Price = items.FirstOrDefault().Price
                                            };

                                            posData.TrnStockInLines.InsertOnSubmit(newStockInLine);

                                            var currentItem = from d in posData.MstItems
                                                              where d.Id == newStockInLine.ItemId
                                                              select d;

                                            if (currentItem.Any())
                                            {
                                                Decimal currentOnHandQuantity = currentItem.FirstOrDefault().OnhandQuantity;
                                                Decimal totalQuantity = currentOnHandQuantity + Convert.ToDecimal(quantity);
                                                var updateItem = currentItem.FirstOrDefault();
                                                updateItem.OnhandQuantity = totalQuantity;
                                            }

                                            posData.SubmitChanges();
                                            Console.WriteLine(item + " was successfuly saved!");
                                        }
                                    }
                                }
                            }
                        }
                    }

                    File.Delete(file);
                }

                if (!hasFiles)
                {
                    Console.WriteLine("No XLS Files Found...");
                    Console.WriteLine();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        // ===========
        // Main Method
        // ===========
        static void Main(string[] args)
        {
            int i = 0;
            string xlsPath = "";
            foreach (var arg in args)
            {
                if (i == 0) { xlsPath = arg; }
                i++;
            }

            Console.WriteLine("Innosoft POS XLS Uploader");
            Console.WriteLine("Version: 1.20170907      ");
            Console.WriteLine("=========================");

            while (true)
            {
                try
                {
                    if (!xlsPath.Equals(""))
                    {
                        Console.WriteLine();
                        SyncStockIn(xlsPath);
                    }
                    else
                    {
                        Console.WriteLine();
                        Console.WriteLine("Please provide XLS path.");
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }

                Thread.Sleep(5000);
            }
        }
    }
}
