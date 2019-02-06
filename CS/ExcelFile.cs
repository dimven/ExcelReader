using System;
using OfficeOpenXml;
using System.IO;
using System.Linq;

//Please refer to the included EPPlus license file if you plan to do anything commercial with that code!

namespace EPPlus.ExcelReader
{
    /// <summary>
    /// EPPlus Excel workbook entry point
    /// </summary>
    public static class ExcelFile
    {
        /// <summary>
        /// Reads the entire content of an excel sheet. 
        /// </summary>
        /// <param name="file">a file object (use File.FromPath)</param>
        /// <param name="sheetName">the sheet's name as a string</param>
        /// <param name="byColumn">reads the data by columns(default) or by rows</param>
        /// <param name="range">you can limit the content you'd like to read, by providing a valid excel address range, such as "A1:C5"</param>
        /// <returns></returns>
        ///<returns name="data">The excel sheet content</returns>
        /// <search>excel read</search>
        public static object[][] Read(FileInfo file, string sheetName, bool byColumn = true, string range = "")
        {
            using (ExcelPackage pck = new ExcelPackage(file))
            {
                var ws = pck.Workbook.Worksheets.FirstOrDefault(x => x.Name == sheetName);
                if (ws == null)
                {
                    throw new IndexOutOfRangeException("No sheet found with such name!");
                }
                
                if (ws.Dimension == null)
                {
                    return new object[][] {new object[] {}};
                }
                
                int colNum, rowNum, colStart, rowStart;
                ExcelRange cells;
                if (String.IsNullOrEmpty(range))
                {
                    colNum = ws.Dimension.Columns;
                    rowNum = ws.Dimension.Rows;
                    colStart = ws.Dimension.Start.Column;
                    rowStart = ws.Dimension.Start.Row;
                    cells = ws.Cells;
                }
                else
                {
                    if (!ExcelCellBase.IsValidAddress(range))
                    {
                        throw new FormatException("That is not a valid excel range! Try a range like A1:C5.");
                    }
                    
                    cells = ws.Cells[range];
                    colNum = cells.Columns;
                    rowNum = cells.Rows;
                    colStart = cells.Start.Column;
                    rowStart = cells.Start.Row;
                }
                
                var data = byColumn ? new object[colNum][] : new object[rowNum][];
                if (byColumn)
                {
                    for (int i = 0; i < colNum; i++)
                    {
                        data[i] = new object[rowNum];
                        for (int j = 0; j < rowNum; j++)
                        {
                            data[i][j] = cells[j + rowStart, i + colStart].Value;
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < rowNum; i++)
                    {
                        data[i] = new object[colNum];
                        for (int j = 0; j < colNum; j++)
                        {
                            data[i][j] = cells[i + rowStart, j + colStart].Value;
                        }
                    }
                }
                return data;
            }
        }
        
        /// <summary>
        /// Gets all of the sheet names stored in this excel workbook file.
        /// </summary>
        /// <param name="file">a file object (use File.FromPath)</param>
        /// <returns>names</returns>
        public static string[] SheetNames(FileInfo file)
        {
            using (ExcelPackage pck = new ExcelPackage(file))
            {
                return pck.Workbook.Worksheets.Select(ws => ws.Name).ToArray();
            }
        }
        
        /// <summary>
        /// Gets all of the named ranges stored in an excel file. If the worksheet parameter is blank,
        /// only ranges registered in the workbook will be fetched.
        /// </summary>
        /// <param name="file">a file object (use File.FromPath)</param>
        /// <param name="sheetName">the name of the sheet as a string</param>
        /// <returns>names</returns>
        public static NamedRange[] NamedRanges(FileInfo file, string sheetName="")
        {
            
        	using (ExcelPackage pck = new ExcelPackage(file))
            {
            	if(String.IsNullOrEmpty(sheetName))
            	   {
            		return pck.Workbook.Names.Select(r => new NamedRange(r.Name, r.Address)).ToArray();
            	   }
            	return pck.Workbook.Worksheets.First(ws => ws.Name == sheetName).Names.Select(r => new NamedRange(r.Name, r.Address)).ToArray();
            }
        }
    }
    
    
    /// <summary>
    /// A placeholder class used to display named ranges nicely in Dynamo
    /// </summary>
    public class NamedRange
        {
        	string name, address;
        	internal NamedRange(string _name, string _addr)
        	{
        		name = _name;
        		address = _addr;
        	}
        	
        	/// <summary>
        	/// standard dynamo preview
        	/// </summary>
        	public override string ToString()
        	{
        		return name;
        	}
        	
        	/// <summary>
        	/// fetches the address of a range as a basic string
        	/// </summary>
        	public object GetRangeAddress()
        	{
        		if (address.Contains(","))
        		{
        			return address.Split(',').Select(a => formatAddress(a)).ToArray();
        		}
        		return formatAddress(address);
        	}
        	
        	private static string formatAddress(string a)
        	{
        		return a.Substring(a.IndexOf('!') + 1).Replace("$", "");
        	}
        }
}
