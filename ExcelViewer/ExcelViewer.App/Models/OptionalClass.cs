using ClosedXML.Excel;
using System;
using static ExcelViewer.App.Controllers.HomeController;

namespace ExcelViewer.App.Models
{



    public static class OptionalClass
    {

        public static string CellExists<TEntity>(IXLCell cell, int cellIndex) where TEntity : class, new()
        {
            TEntity obj = new TEntity();
            if (obj is Product)
            {
                if (cell.Value.ToString() is null || cell.Value.ToString() == "")
                {
                    return "gray";
                }
                else
                {
                    if (cellIndex == 0)//Name
                    {
                        if (cell.Value.ToString().Length >= 5 && cell.Value.ToString().Length <= 20)
                        {
                            return "green";
                        }
                        else
                        {
                            return "red";
                        }
                    }
                    else if (cellIndex == 1)//Description
                    {
                        if (cell.Value.ToString().Length <= 100)
                        {
                            return "green";
                        }
                        else
                        {
                            return "red";
                        }
                    }
                    else if (cellIndex == 2)//Price
                    {                     
                        if (decimal.TryParse((cell.Value.ToString()), out decimal result))
                        {
                            return "green";
                        }
                        else
                        {
                            return "red";
                        }
                    }
                    else if (cellIndex == 3)//Serial Number
                    {
                        if (long.TryParse((cell.Value.ToString()), out long result))
                        {
                            return "green";
                        }
                        else
                        {
                            return "red";
                        }
                    }
                    else if (cellIndex == 4)//Supplier
                    {
                        if (cell.Value.ToString().Length >= 1)
                        {
                            return "green";
                        }
                        else
                        {
                            return "red";
                        }
                    }
                }

            }
            return "";
        }

    }
}
