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
        private static POS13db.POS13dbDataContext posData;

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

        // =============
        // Sync Stock In
        // =============
        public static void SyncStockInAndStockOut(String xlsPath, String database)
        {
            try
            {
                var newConnectionString = "Data Source=localhost;Initial Catalog=" + database + ";Integrated Security=True";
                posData = new POS13db.POS13dbDataContext(newConnectionString);

                List<String> files = new List<String>(Directory.EnumerateFiles(xlsPath));
                foreach (var file in files)
                {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                    Excel.Range range = xlWorkSheet.UsedRange;

                    Int32 stockInId = 0;
                    Int32 stockOutId = 0;

                    Int32 row = range.Rows.Count;
                    Int32 column = range.Columns.Count;

                    Console.WriteLine("Saving " + (range.Cells[2, 1] as Excel.Range).Value2 + "...");

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

                                // ========
                                // Defaults 
                                // ========
                                var defaultPeriod = from d in posData.MstPeriods
                                                    select d;

                                var defaultSettings = from d in posData.SysSettings
                                                      select d;

                                var defaultExpenseAccounts = from d in posData.MstAccounts
                                                             where d.AccountType.Equals("EXPENSES")
                                                             select d;

                                // =============
                                // Sync Stock In
                                // =============
                                var stockIn = from d in posData.TrnStockIns
                                              where d.Remarks.Equals(documentReference)
                                              select d;

                                if (!stockIn.Any())
                                {
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
                                            Quantity = Convert.ToDecimal(quantity),
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

                                        Console.WriteLine("Inserting Stock-In Item: " + item + "...");

                                        posData.SubmitChanges();
                                        Console.WriteLine(item + " was successfuly saved!");
                                    }
                                }

                                // ==============
                                // Sync Stock Out
                                // ==============
                                var stockOut = from d in posData.TrnStockOuts
                                               where d.Remarks.Equals(documentReference)
                                               select d;

                                if (!stockOut.Any())
                                {
                                    var defaultStockOutNumber = defaultPeriod.FirstOrDefault().Period + "-000001";

                                    var lastStockOut = from d in posData.TrnStockOuts.OrderByDescending(d => d.Id)
                                                       select d;

                                    if (lastStockOut.Any())
                                    {
                                        var lastStockOutNumber = lastStockOut.FirstOrDefault().StockOutNumber.ToString();
                                        int secondIndex = lastStockOutNumber.IndexOf('-', lastStockOutNumber.IndexOf('-'));
                                        var stockOutNumberSplitStringValue = lastStockOutNumber.Substring(secondIndex + 1);
                                        var stockOutNumber = Convert.ToInt32(stockOutNumberSplitStringValue) + 000001;
                                        defaultStockOutNumber = defaultPeriod.FirstOrDefault().Period + "-" + FillLeadingZeroes(stockOutNumber, 6);
                                    }

                                    POS13db.TrnStockOut newStockOut = new POS13db.TrnStockOut
                                    {
                                        PeriodId = defaultPeriod.FirstOrDefault().Id,
                                        StockOutDate = DateTime.Today,
                                        StockOutNumber = defaultStockOutNumber,
                                        AccountId = defaultExpenseAccounts.FirstOrDefault().Id,
                                        Remarks = documentReference,
                                        PreparedBy = defaultSettings.FirstOrDefault().PostUserId,
                                        CheckedBy = defaultSettings.FirstOrDefault().PostUserId,
                                        ApprovedBy = defaultSettings.FirstOrDefault().PostUserId,
                                        IsLocked = true,
                                        EntryUserId = defaultSettings.FirstOrDefault().PostUserId,
                                        EntryDateTime = DateTime.Now,
                                        UpdateUserId = defaultSettings.FirstOrDefault().PostUserId,
                                        UpdateDateTime = DateTime.Now
                                    };

                                    posData.TrnStockOuts.InsertOnSubmit(newStockOut);
                                    posData.SubmitChanges();

                                    stockOutId = newStockOut.Id;

                                    var items = from d in posData.MstItems
                                                where d.BarCode.Equals(barCode)
                                                select d;

                                    if (items.Any())
                                    {
                                        POS13db.TrnStockOutLine newStockOutLine = new POS13db.TrnStockOutLine
                                        {
                                            StockOutId = stockOutId,
                                            ItemId = items.FirstOrDefault().Id,
                                            UnitId = items.FirstOrDefault().UnitId,
                                            Quantity = Convert.ToDecimal(quantity),
                                            Cost = Convert.ToDecimal(cost),
                                            Amount = Convert.ToDecimal(amount),
                                            AssetAccountId = items.FirstOrDefault().AssetAccountId
                                        };

                                        posData.TrnStockOutLines.InsertOnSubmit(newStockOutLine);

                                        var currentItem = from d in posData.MstItems
                                                          where d.Id == newStockOutLine.ItemId
                                                          select d;

                                        if (currentItem.Any())
                                        {
                                            Decimal currentOnHandQuantity = currentItem.FirstOrDefault().OnhandQuantity;
                                            Decimal totalQuantity = currentOnHandQuantity - Convert.ToDecimal(quantity);
                                            var updateItem = currentItem.FirstOrDefault();
                                            updateItem.OnhandQuantity = totalQuantity;
                                        }

                                        Console.WriteLine("Inserting Stock-Out Item: " + item + "...");

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

                                    // ==============
                                    // Stock In Items
                                    // ==============
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
                                                Quantity = Convert.ToDecimal(quantity),
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

                                            Console.WriteLine("Inserting Stock-In Item: " + item + "...");

                                            posData.SubmitChanges();
                                            Console.WriteLine(item + " was successfuly saved!");
                                        }
                                    }

                                    // ===============
                                    // Stock Out Items
                                    // ===============
                                    if (stockOutId > 0)
                                    {
                                        var items = from d in posData.MstItems
                                                    where d.BarCode.Equals(barCode)
                                                    select d;

                                        if (items.Any())
                                        {
                                            POS13db.TrnStockOutLine newStockOutLine = new POS13db.TrnStockOutLine
                                            {
                                                StockOutId = stockOutId,
                                                ItemId = items.FirstOrDefault().Id,
                                                UnitId = items.FirstOrDefault().UnitId,
                                                Quantity = Convert.ToDecimal(quantity),
                                                Cost = Convert.ToDecimal(cost),
                                                Amount = Convert.ToDecimal(amount),
                                                AssetAccountId = items.FirstOrDefault().AssetAccountId
                                            };

                                            posData.TrnStockOutLines.InsertOnSubmit(newStockOutLine);

                                            var currentItem = from d in posData.MstItems
                                                              where d.Id == newStockOutLine.ItemId
                                                              select d;

                                            if (currentItem.Any())
                                            {
                                                Decimal currentOnHandQuantity = currentItem.FirstOrDefault().OnhandQuantity;
                                                Decimal totalQuantity = currentOnHandQuantity - Convert.ToDecimal(quantity);
                                                var updateItem = currentItem.FirstOrDefault();
                                                updateItem.OnhandQuantity = totalQuantity;
                                            }

                                            Console.WriteLine("Inserting Stock-Out Item: " + item + "...");

                                            posData.SubmitChanges();
                                            Console.WriteLine(item + " was successfuly saved!");
                                        }
                                    }
                                }
                            }
                        }

                        Console.WriteLine();
                    }

                    File.Delete(file);
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
            String xlsPath = "", database = "";

            int i = 0;
            foreach (var arg in args)
            {
                if (i == 0) { xlsPath = arg; }
                else if (i == 1) { database = arg; }
                i++;
            }

            Console.WriteLine("====================================");
            Console.WriteLine("POS XLS Uploader Version: 1.20180116");
            Console.WriteLine("====================================");

            while (true)
            {
                try
                {
                    if (!xlsPath.Equals(""))
                    {
                        Console.WriteLine();
                        Console.WriteLine("Scanning XLS Files...");

                        List<string> files = new List<string>(Directory.EnumerateFiles(xlsPath));
                        if (files.Any())
                        {
                            SyncStockInAndStockOut(xlsPath, database);
                        }
                        else
                        {
                            Console.WriteLine();
                            Console.WriteLine("No XLS Found...");
                        }
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
