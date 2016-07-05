//Created By Himanshu Saraowgi

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ArtOfTest.WebAii.Controls.HtmlControls;
using ArtOfTest.WebAii.Core;
using ArtOfTest.Common;
using ArtOfTest.WebAii.Design.Execution;




namespace CustomFunc
{

    public class Func
    {
        /// <summary>
        /// This Function reads data from excel by inputing the String FilePath,Integer column, Integer Row, String Sheetname,Boolean Visibility of excel
        /// </summary>
        /// <param name="filePath">Enter String Full File Path</param>
        /// <param name="Col">Enter the Integer Column Number(Cannot exceed 24 i.e.Z)</param>
        /// <param name="Row">Enter the Integer Row Number</param>
        /// <param name="sheetName">Enter String sheetname</param>
        /// <param name="Visible">Enter true if you want the excel visible during run,esle enter false</param>
        /// <returns>String value from excel</returns>

        public static String DataFromExcel(String filePath, int Col, int Row, String sheetName, Boolean Visible)
        {
            Excel.Application myExcelApp;
            Excel.Workbook myExcelWorkbook;
            myExcelApp = new Excel.Application();
            myExcelApp.Visible = Visible;
            String fileName = filePath;
            myExcelWorkbook = myExcelApp.Workbooks.Open(fileName);
            Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.Sheets[sheetName];
            String Row1;
            Row1 = Row.ToString();
            //string Col1 = char.ConvertFromUtf32(Col + 64);
            string Col1 = GetExcelColumnName(Col);
            String data = myExcelWorksheet.get_Range(Col1 + Row1).Value.ToString();
            string tmpName = Path.GetTempFileName();
            File.Delete(tmpName);
            myExcelWorkbook.SaveAs(tmpName);
            myExcelWorkbook.Close(false);
            myExcelApp.Quit();
            File.Delete(filePath);
            File.Move(tmpName, filePath);
            //Marshal.ReleaseComObject(myExcelWorksheet);
            //Marshal.ReleaseComObject(myExcelWorkbook);
            //Marshal.ReleaseComObject(myExcelApp);
            return data;

        }
        /// <summary>
        /// This Function writes data to excel by inputing the String FilePath,Integer column, Integer Row, String Sheetname,Boolean Visibility of excel,String Data
        /// </summary>
        /// <param name="filePath">Enter String Full File Path</param>
        /// <param name="Col">Enter the Integer Column Number(Cannot exceed 24 i.e.Z)</param>
        /// <param name="Row">Enter the Integer Row Number</param>
        /// <param name="sheetName">Enter String sheetname</param>
        /// <param name="Visible">Enter true if you want the excel visible during run,esle enter false</param>
        /// <param name="Data">Enter the String value you want to add to excel</param>
        /// <returns></returns>

        public static void DataToExcel(String filePath, int Col, int Row, String sheetName, Boolean Visible, String Data)
        {
            Excel.Application myExcelApp;
            Excel.Workbook myExcelWorkbook;
            myExcelApp = new Excel.Application();
            myExcelApp.Visible = Visible;
            String fileName = filePath;
            myExcelWorkbook = myExcelApp.Workbooks.Open(fileName);
            Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.Sheets[sheetName];
            String Row1;
            Row1 = Row.ToString();
            //string Col1 = char.ConvertFromUtf32(Col + 64);
            string Col1 = GetExcelColumnName(Col);
            myExcelWorksheet.get_Range(Col1 + Row1).Formula = Data;

            string tmpName = Path.GetTempFileName();
            File.Delete(tmpName);
            myExcelWorkbook.SaveAs(tmpName);
            myExcelWorkbook.Close(false);
            myExcelApp.Quit();
            File.Delete(filePath);
            File.Move(tmpName, filePath);
            //Marshal.ReleaseComObject(myExcelWorksheet);
            //Marshal.ReleaseComObject(myExcelWorkbook);
            //Marshal.ReleaseComObject(myExcelApp);

        }


        /// <summary>
        /// This Function converts a column number to String column name as depicted in Excel(A,B,...AA,AB....BA,BB)
        /// </summary>
        /// <param name="columnnumber">Enter yhe column number you want to convert</param>
        /// <returns>String value of Column name</returns>

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }



        /// <summary>
        /// This Function returns the row count in excel by inputing the String FilePath,String Sheetname
        /// </summary>
        /// <param name="filePath">Enter String Full File Path</param>
        /// <param name="sheetName">Enter String sheetname</param>
        /// <returns>Row number</returns>

        public static int RowCount(String filePath, String sheetName)
        {
            Excel.Application myExcelApp;
            Excel.Workbook myExcelWorkbook;

            myExcelApp = new Excel.Application();
            myExcelApp.Visible = false;
            String fileName = filePath;
            myExcelWorkbook = myExcelApp.Workbooks.Open(fileName);
            Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.Sheets[sheetName];
            

            int Row = myExcelWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            //Func.RowCol.Col = myExcelWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            myExcelWorkbook.Close(false);
            myExcelApp.Quit();
            //Marshal.ReleaseComObject(myExcelWorksheet);
            //Marshal.ReleaseComObject(myExcelWorkbook);
            //Marshal.ReleaseComObject(myExcelApp);
            return Row;
        }

        /// <summary>
        /// This Function returns the column count in excel by inputing the String FilePath,String Sheetname
        /// </summary>
        /// <param name="filePath">Enter String Full File Path</param>
        /// <param name="sheetName">Enter String sheetname</param>
        /// <returns>Column count</returns>

        public static int ColCount(String filePath, String sheetName)
        {
            Excel.Application myExcelApp;
            Excel.Workbook myExcelWorkbook;

            myExcelApp = new Excel.Application();
            myExcelApp.Visible = false;
            String fileName = filePath;
            myExcelWorkbook = myExcelApp.Workbooks.Open(fileName);
            Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.Sheets[sheetName];


            //Func.RowCol.Row = myExcelWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int Col = myExcelWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            myExcelWorkbook.Close(false);
            myExcelApp.Quit();
            //Marshal.ReleaseComObject(myExcelWorksheet);
            //Marshal.ReleaseComObject(myExcelWorkbook);
            //Marshal.ReleaseComObject(myExcelApp);
            return Col;
        }

        /// <summary>
        /// Create msgbox with custom text
        /// </summary>
        /// <param name="message">Message to be passed to the messagebox</param>
        public static void MsgBox(String message)
        {
            MessageBox.Show(message);
        }

        /// <summary>
        /// This Function adds options of Double, right clicking to the fuction "ClickFromTableByText" and also an option to scroll to the desired element before click.
        /// </summary>
        /// <param name="table">Pass the Htmltable object</param>
        /// <param name="text">Pass the string you are looking for in the table</param>
        /// <param name="Column">Pass the column number of the table which has the button</param>
        /// <param name="click">Enter Double for double click, Single for single click, Right for right click </param>
        /// /// <param name="scroll">Enter true if you want to scroll to the element, else enter false </param>
        /// <returns>Tru if value was found and clicked, else false</returns>
        public static Boolean ClickFromTableByText(HtmlTable table, String Text, int Column, string click, Boolean scroll)
        {
            Boolean RetunValue = false;
            foreach (HtmlTableRow row in table.AllRows)
            {

                foreach (HtmlTableCell cell in row.Cells)
                {
                    
                    if(cell.ChildNodes.Count>=1)
                    {
                        if (cell.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).TextContent.ToString() == Text)
                            {
                                
                                   RetunValue = true;
                                   HtmlTableCell cell1 = row.Cells[Column];
                                   if (scroll == true)
                                       cell1.ScrollToVisible();
                                   if (click == "Double")
                                       cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftDoubleClick);
                                   else if (click == "Single")
                                       cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                                   else if (click == "Right")
                                       cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.RightClick);
                                   else
                                       throw new System.Exception("Mouse Click Type wrong");
                                   break;
                                    
                                

                            }
                        }
                        if (cell.ChildNodes.ElementAt(0).InnerText != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).InnerText == Text)
                            {
                                RetunValue = true;
                                HtmlTableCell cell1 = row.Cells[Column];
                                if (scroll == true)
                                    cell1.ScrollToVisible();
                                if (click == "Double")
                                    cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftDoubleClick);
                                else if (click == "Single")
                                    cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                                else if (click == "Right")
                                    cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.RightClick);
                                else
                                    throw new System.Exception("Mouse Click Type wrong");
                                break;

                            }
                        }
                    }
                    else
                    {
                        if (cell.TextContent.ToString() == Text)
                        {
                            RetunValue = true;
                            HtmlTableCell cell1 = row.Cells[Column];
                            if(scroll==true)
                                cell1.ScrollToVisible();
                            if (click == "Double")
                                cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftDoubleClick);
                            else if (click == "Single")
                                cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                            else if (click == "Right")
                                cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.RightClick);
                            else
                                throw new System.Exception("Mouse Click Type wrong");
                            
                            break;
                                                      

                        }
                    }
                }
                if (RetunValue == true)
                {
                    break;
                }

            }
            return RetunValue;
        }

        /// <summary>
        /// This Function returns the cell from table from a specified column and the row which has the specified text..
        /// </summary>
        /// <param name="table">Pass the Htmltable object</param>
        /// <param name="text">Pass the string you are looking for in the table</param>
        /// <param name="Column">Pass the column number of the table which has the button</param>
        /// <returns>The HTMLTable Cell required</returns>
        public static HtmlTableCell CellFromTableByText(HtmlTable table, String Text, int Column)
        {
            HtmlTableCell RetunValue = null;
            foreach (HtmlTableRow row in table.AllRows)
            {

                foreach (HtmlTableCell cell in row.Cells)
                {

                    if (cell.ChildNodes.Count >= 1)
                    {
                        if (cell.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).TextContent.ToString() == Text)
                            {
                                
                                HtmlTableCell cell1 = row.Cells[Column];
                                RetunValue = cell1;
                                
                                break;

                            }
                        }
                        if (cell.ChildNodes.ElementAt(0).InnerText != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).InnerText == Text)
                            {
                                
                                HtmlTableCell cell1 = row.Cells[Column];
                                RetunValue = cell1;
                                break;

                            }
                        }
                    }
                    else
                    {
                        if (cell.TextContent.ToString() == Text)
                        {
                            
                            HtmlTableCell cell1 = row.Cells[Column];
                            RetunValue = cell1;
                            break;


                        }
                    }
                }
                if (RetunValue != null)
                {
                    break;
                }

            }
            return RetunValue;
        }

        /// <summary>
        /// This Function finds a value in a table and return true if found.
        /// </summary>
        /// <param name="table">Pass the Htmltable object</param>
        /// <param name="text">Pass the string you are looking for in the table</param>
        /// <param name="Column">Pass the column number of the table which has the button</param>
        /// <returns>true if value found, else false</returns>
        public static Boolean FindTextInTable(HtmlTable table, String Text)
        {
            Boolean RetunValue = false;
            foreach (HtmlTableRow row in table.AllRows)
            {

                foreach (HtmlTableCell cell in row.Cells)
                {
                    if (cell.ChildNodes.Count >= 1)
                    {
                        if (cell.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).TextContent.ToString() == Text)
                            {
                                RetunValue = true;
                                break;
                            }
                        }
                        if (cell.ChildNodes.ElementAt(0).InnerText != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).InnerText == Text)
                            {
                                RetunValue = true;
                                break;
                            }
                        }

                    }
                    else
                    {
                        if (cell.TextContent.ToString() == Text)
                        {
                            RetunValue = true;
                            break;
                        }
                    }
                }
                if (RetunValue == true)
                {
                    break;
                }

            }
            return RetunValue;
        }

        /// <summary>
        /// This Function Single Clicks on a button in a Html table according to a text entered.
        /// </summary>
        /// <param name="table">Pass the Htmltable object</param>
        /// <param name="text">Pass the string you are looking for in the table</param>
        /// <param name="Column">Pass the column number of the table which has the button</param>
        /// <returns>Tru if value was found and clicked, else false</returns>
        public static Boolean ClickFromTableByText(HtmlTable table, String Text, int Column)
        {
            Boolean RetunValue = false;
            foreach (HtmlTableRow row in table.AllRows)
            {

                foreach (HtmlTableCell cell in row.Cells)
                {
                    if (cell.ChildNodes.Count >= 1)
                    {
                        if (cell.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).TextContent.ToString() == Text)
                            {
                                RetunValue = true;
                                HtmlTableCell cell1 = row.Cells[Column];

                                cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                                break;

                            }
                        }
                        if (cell.ChildNodes.ElementAt(0).InnerText != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).InnerText == Text)
                            {
                                RetunValue = true;
                                HtmlTableCell cell1 = row.Cells[Column];

                                cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                                break;

                            }
                        }

                    }
                    else
                    {
                        if (cell.TextContent.ToString() == Text)
                        {
                            RetunValue = true;
                            HtmlTableCell cell1 = row.Cells[Column];
                            //cell1.ScrollToVisible();
                            cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                            break;
                            

                        }
                    }
                }
                if (RetunValue == true)
                {
                    break;
                }

            }
            return RetunValue;
        }
        /// <summary>
        /// This Function performs a key action independent of object
        /// </summary>
        /// <param name="key">Pass the key you want to press-"Down" for down arrow key, "Enter" for Enter key,Esc,Up,Right,Left</param>
        public static void KeyAction(String Key)
        {
            Desktop DesktopObject = new Desktop();
            if (Key == "Down")
                DesktopObject.KeyBoard.KeyPress(Keys.Down);
                
            if (Key == "Enter")
                DesktopObject.KeyBoard.KeyPress(Keys.Enter);
            if (Key == "Esc")
                DesktopObject.KeyBoard.KeyPress(Keys.Escape);
            if (Key == "Left")
                DesktopObject.KeyBoard.KeyPress(Keys.Left);
            if (Key == "Right")
                DesktopObject.KeyBoard.KeyPress(Keys.Right);
            if (Key == "Up")
                DesktopObject.KeyBoard.KeyPress(Keys.Up);
            
        }

        /// <summary>
        /// This Function returns the value on a column for the row selected by finding a string in the table.
        /// </summary>
        /// <param name="table">Pass the Htmltable object</param>
        /// <param name="text">Pass the string you are looking for in the table</param>
        /// <param name="Column">Pass the column number of the table which has the value we are looking for</param>
        /// <returns>Value on the required column</returns>
        public static String GetDataFromTableByText(HtmlTable table, String Text, int Column)
        {
            String RetunValue = "";
            Boolean flag = false;
            foreach (HtmlTableRow row in table.AllRows)
            {

                foreach (HtmlTableCell cell in row.Cells)
                {
                    if (cell.ChildNodes.Count >= 1)
                    {
                        if (cell.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).TextContent.ToString() == Text)
                            {
                                flag = true;
                                HtmlTableCell cell1 = row.Cells[Column];

                                if (cell1.ChildNodes.Count >= 1)
                                {
                                    if (cell1.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                                        RetunValue = cell1.ChildNodes.ElementAt(0).TextContent.ToString();
                                    else
                                        RetunValue = cell1.ChildNodes.ElementAt(0).InnerText.ToString();
                                }
                                else
                                {
                                    if (cell1.TextContent.ToString() != "")
                                        RetunValue = cell1.TextContent.ToString();
                                }                                                                                          
        
                                
                                break;

                            }
                        }
                        if (cell.ChildNodes.ElementAt(0).InnerText != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).InnerText == Text)
                            {
                                flag = true;
                                HtmlTableCell cell1 = row.Cells[Column];

                                if (cell1.ChildNodes.Count >= 1)
                                {
                                    if (cell1.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                                        RetunValue = cell1.ChildNodes.ElementAt(0).TextContent.ToString();
                                    else
                                        RetunValue = cell1.ChildNodes.ElementAt(0).InnerText.ToString();
                                }
                                else
                                {
                                    if (cell1.TextContent.ToString() != "")
                                        RetunValue = cell1.TextContent.ToString();
                                }                                                                                          
        
                                break;

                            }
                        }

                    }
                    else
                    {
                        if (cell.TextContent.ToString() == Text)
                        {
                            flag = true;
                            HtmlTableCell cell1 = row.Cells[Column];
                            //cell1.ScrollToVisible();
                            if (cell1.ChildNodes.Count >= 1)
                            {
                                if (cell1.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                                    RetunValue = cell1.ChildNodes.ElementAt(0).TextContent.ToString();
                                else
                                    RetunValue = cell1.ChildNodes.ElementAt(0).InnerText.ToString();
                            }
                            else
                            {
                                if (cell1.TextContent.ToString() != "")
                                    RetunValue = cell1.TextContent.ToString();
                            }                                                                                          
        
                            break;


                        }
                    }
                }
                if (flag == true)
                {
                    break;
                }

            }
            return RetunValue;
        }


        /// <summary>
        /// This Function writes data to textFile by inputing the String FilePath,String Data
        /// </summary>
        /// <param name="filePath">Enter String Full File Path</param>
        /// <param name="Data">Enter the String value you want to add to txt file</param>
        /// <returns></returns>

        public static void DataToTxt(String filePath,String Data)
        {
            StreamWriter myfile = new StreamWriter(filePath);
            myfile.WriteLine(Data);
            myfile.Close();

        }



        /// <summary>
        /// This Function reads data from textFile by inputing the String FilePath
        /// </summary>
        /// <param name="filePath">Enter String Full File Path</param>
        /// <returns></returns>

        public static String DataFromTxt(String filePath)
        {
            StreamReader myfile = new StreamReader(filePath);
            String text = myfile.ReadLine();
            myfile.Close();
            return text;
        }


        /// <summary>
        /// This function gets the data in string form from the row and column of the csv file referred.
        /// </summary>
        /// <param name="filename">String file path of the csv file</param>
        /// <param name="rownum">row number in integer</param>
        /// <param name="colnum">Column number in Integer</param>
        /// <returns>String value of data on the row and column intersection of the csv file.</returns>
        public static String getCSVdata(String filename, int rownum, int colnum, char delimiter)
        {
            var reader = new StreamReader(File.OpenRead(@filename));
            int runningRow = 0;
            while (!reader.EndOfStream)
            {
                runningRow++;
                var line = reader.ReadLine();
                if (runningRow == rownum)
                {
                    var values = line.Split(delimiter);
                    reader.Close();
                    return values[colnum - 1];
                }

            }
            reader.Close();
            return "";
        }

        /// <summary>
        /// This method reurns an integer value of the number of rows in the csv file.
        /// </summary>
        /// <param name="filename">String file path of the csv file</param>
        /// <returns>Integer value of number of rows in the csv</returns>
        public static int getCSVrowCount(String filename)
        {
            var reader = new StreamReader(File.OpenRead(@filename));
            int runningRow = 0;
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                runningRow++;

            }
            reader.Close();
            return runningRow;
        }
        /// <summary>
        /// This method reurns an integer value of the number of columns in the csv file.
        /// </summary>
        /// <param name="filename">String file path of the csv file</param>
        /// <param name="delimiter">Enter the delimiting character</param>
        /// <returns>Integer value of number of columns in the csv</returns>
        public static int getCSVcolCount(String filename, char delimiter)
        {
            var reader = new StreamReader(File.OpenRead(@filename));
            int runningRow = 0;
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                runningRow++;
                if (runningRow == 1)
                {
                    var values = line.Split(delimiter);
                    reader.Close();
                    return values.Count();
                }

            }
            reader.Close();
            return 0;
        }

        /// <summary>
        /// This method replaces the data at a row column intersection in the csv file with new data as passed to the function or if the row number mentioned is higher than total rows, it adds a row with desired data
        /// </summary>
        /// <param name="filename">String file path of the csv file</param>
        /// <param name="rownum">Integer row number</param>
        /// <param name="colnum">Integer Column Number</param>
        /// <param name="delimiter">Enter the delimiting character</param>
        /// <param name="data">Data to input to the csv</param>
        public static void setCSVdata(String filename, int rownum, int colnum, String data, char delimiter)
        {
            if ((rownum <= getCSVrowCount(filename)) && (colnum <= getCSVcolCount(filename, delimiter)))
            {
                var reader = new StreamReader(File.OpenRead(@filename));
                String wholetext = File.ReadAllText(filename);
                int runningRow = 0;
                string line1 = "";
                var line = "";
                String line2 = " ";
                while (!reader.EndOfStream)
                {
                    runningRow++;
                    line = reader.ReadLine();
                    if (runningRow == rownum)
                    {
                        var values = line.Split(delimiter);

                        //for (int a = 0; a < 4; a++)
                        //listA.Add(values[0]);
                        //listB.Add(values[1]);
                        //
                        values[colnum - 1] = data;



                        line2 = line;
                        line1 = values[0];
                        for (int i = 1; i < values.Count(); i++)
                            line1 = line1 + delimiter + values[i];
                    }


                }
                reader.Close();
                StreamWriter myfile = new StreamWriter(filename);
                //myfile.WriteLine(Data);


                myfile.Write(wholetext.Replace(line2, line1));
                myfile.Close();
            }
            else if ((rownum > getCSVrowCount(filename)) && (colnum <= getCSVcolCount(filename, delimiter)))
            {
                var reader = new StreamReader(File.OpenRead(@filename));
                String wholetext = File.ReadAllText(filename);
                String newrow = "\r\n";
                for (int i = 1; i <= getCSVcolCount(filename, delimiter); i++)
                {
                    if (i == colnum)
                        newrow = newrow + data;
                    else
                        newrow = newrow + "";
                    if (i < getCSVcolCount(filename, delimiter))
                        newrow = newrow + delimiter;
                }

                wholetext = wholetext + newrow;
                reader.Close();
                StreamWriter myfile = new StreamWriter(filename);
                myfile.Write(wholetext);
                myfile.Close();

            }
            else
            {
                throw new System.IndexOutOfRangeException("Column selected does not exist");
            }


        }


        /// <summary>
        /// Override to DataFromExcel() to cater Relative file path
        /// This Function reads data from excel by inputing the relative FilePath,Integer column, Integer Row, String Sheetname, and testContext
        /// 
        /// </summary>
        /// <param name="filePath">Enter path of file relative to the project directory</param>
        /// <param name="Col">Enter the Integer Column Number(Cannot exceed 24 i.e.Z)</param>
        /// <param name="Row">Enter the Integer Row Number</param>
        /// <param name="sheetName">Enter String sheetname</param>
        /// <param name="TestContext">Enter 'this' to use relative path </param>
        /// <returns>String value from excel</returns>

        public static String DataFromExcel(String filePath, int Col, int Row, String sheetName, ExecutionContext TestContext)
        {
            Excel.Application myExcelApp;
            Excel.Workbook myExcelWorkbook;
            myExcelApp = new Excel.Application();
            myExcelApp.Visible = false;
            String fileName = TestContext.DeploymentDirectory+"\\"+filePath;
            myExcelWorkbook = myExcelApp.Workbooks.Open(fileName);
            Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.Sheets[sheetName];
            String Row1;
            Row1 = Row.ToString();
            //string Col1 = char.ConvertFromUtf32(Col + 64);
            string Col1 = GetExcelColumnName(Col);
            String data = myExcelWorksheet.get_Range(Col1 + Row1).Value.ToString();
            string tmpName = Path.GetTempFileName();
            File.Delete(tmpName);
            myExcelWorkbook.SaveAs(tmpName);
            myExcelWorkbook.Close(false);
            myExcelApp.Quit();
            File.Delete(fileName);
            File.Move(tmpName, fileName);
            //Marshal.ReleaseComObject(myExcelWorksheet);
            //Marshal.ReleaseComObject(myExcelWorkbook);
            //Marshal.ReleaseComObject(myExcelApp);
            return data;

        }
        /// <summary>
        /// Override to DataToExcel() to cater Relative file path
        /// This Function writes data to excel by inputing the relative String FilePath,Integer column, Integer Row, String Sheetname,String Data and text context 
        /// </summary>
        /// <param name="filePath">Enter String Full File Path</param>
        /// <param name="Col">Enter the Integer Column Number(Cannot exceed 24 i.e.Z)</param>
        /// <param name="Row">Enter the Integer Row Number</param>
        /// <param name="sheetName">Enter String sheetname</param>
        /// <param name="TestContext">Enter 'this' to use relative path</param>
        /// <param name="Data">Enter the String value you want to add to excel</param>
        /// <returns></returns>

        public static void DataToExcel(String filePath, int Col, int Row, String sheetName, ExecutionContext TestContext, String Data)
        {
            Excel.Application myExcelApp;
            Excel.Workbook myExcelWorkbook;
            myExcelApp = new Excel.Application();
            myExcelApp.Visible = false;
            
            String fileName = TestContext.DeploymentDirectory +"\\"+filePath;
             
            myExcelWorkbook = myExcelApp.Workbooks.Open(fileName);
            
            Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.Sheets[sheetName];
            String Row1;
            Row1 = Row.ToString();
            //string Col1 = char.ConvertFromUtf32(Col + 64);
            string Col1 = GetExcelColumnName(Col);
            myExcelWorksheet.get_Range(Col1 + Row1).Formula = Data;

            string tmpName = Path.GetTempFileName();
            File.Delete(tmpName);
            myExcelWorkbook.SaveAs(tmpName);
            myExcelWorkbook.Close(false);
            myExcelApp.Quit();
            File.Delete(fileName);
            File.Move(tmpName, fileName);
            //Marshal.ReleaseComObject(myExcelWorksheet);
            //Marshal.ReleaseComObject(myExcelWorkbook);
            //Marshal.ReleaseComObject(myExcelApp);

        }

        /// <summary>
        /// Override to RowCount() to cater Relative file path
        /// This Function returns the row count in excel by inputing the String FilePath,String Sheetname
        /// </summary>
        /// <param name="filePath">Enter String Full File Path</param>
        /// <param name="sheetName">Enter String sheetname</param>
        /// <param name="testContext">Enter 'this' to use relative path</param>
        /// <returns>Row number</returns>

        public static int RowCount(String filePath, String sheetName,ExecutionContext TestContext)
        {
            Excel.Application myExcelApp;
            Excel.Workbook myExcelWorkbook;

            myExcelApp = new Excel.Application();
            myExcelApp.Visible = false;
            String fileName = TestContext.DeploymentDirectory + "\\" + filePath;
            myExcelWorkbook = myExcelApp.Workbooks.Open(fileName);
            Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.Sheets[sheetName];


            int Row = myExcelWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            //Func.RowCol.Col = myExcelWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            myExcelWorkbook.Close(false);
            myExcelApp.Quit();
            //Marshal.ReleaseComObject(myExcelWorksheet);
            //Marshal.ReleaseComObject(myExcelWorkbook);
            //Marshal.ReleaseComObject(myExcelApp);
            return Row;
        }

        /// <summary>
        /// Override to ColCount() to cater Relative file path
        /// This Function returns the column count in excel by inputing the String FilePath,String Sheetname
        /// </summary>
        /// <param name="filePath">Enter String Full File Path</param>
        /// <param name="sheetName">Enter String sheetname</param>
        /// <param name="testContext">Enter 'this' to use relative path</param>
        /// <returns>Column count</returns>

        public static int ColCount(String filePath, String sheetName,ExecutionContext TestContext)
        {
            Excel.Application myExcelApp;
            Excel.Workbook myExcelWorkbook;

            myExcelApp = new Excel.Application();
            myExcelApp.Visible = false;
            String fileName = TestContext.DeploymentDirectory + "\\" + filePath;
            myExcelWorkbook = myExcelApp.Workbooks.Open(fileName);
            Excel.Worksheet myExcelWorksheet = (Excel.Worksheet)myExcelWorkbook.Sheets[sheetName];


            //Func.RowCol.Row = myExcelWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int Col = myExcelWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;
            myExcelWorkbook.Close(false);
            myExcelApp.Quit();
            //Marshal.ReleaseComObject(myExcelWorksheet);
            //Marshal.ReleaseComObject(myExcelWorkbook);
            //Marshal.ReleaseComObject(myExcelApp);
            return Col;
        }
        /// <summary>
        /// This Function adds options of Double, right clicking to the fuction "ClickFromTableByMultiText" and also an option to scroll to the desired element before click.
        /// </summary>
        /// <param name="table">Pass the Htmltable object</param>
        /// <param name="text">Pass the string you are looking for in the table</param>
        /// <param name="text2">Pass the 2nd string you are looking for in the table</param>
        /// <param name="Column">Pass the column number of the table which has the button</param>
        /// <param name="click">Enter Double for double click, Single for single click, Right for right click </param>
        /// /// <param name="scroll">Enter true if you want to scroll to the element, else enter false </param>
        /// <returns>Tru if value was found and clicked, else false</returns>
        public static Boolean ClickFromTableByMultiText(HtmlTable table, String Text, String Text2, int Column, string click, Boolean scroll)
        {
            Boolean RetunValue = false;
            foreach (HtmlTableRow row in table.AllRows)
            {

                foreach (HtmlTableCell cell in row.Cells)
                {

                    if (cell.ChildNodes.Count >= 1)
                    {
                        if (cell.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).TextContent.ToString() == Text)
                            {
                                foreach (HtmlTableCell cell2 in row.Cells)
                                {
                                    if (cell2.ChildNodes.ElementAt(0).TextContent.ToString() == Text2)
                                    {
                                        RetunValue = true;
                                        HtmlTableCell cell1 = row.Cells[Column];
                                        if (scroll == true)
                                            cell1.ScrollToVisible();
                                        if (click == "Double")
                                            cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftDoubleClick);
                                        else if (click == "Single")
                                            cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                                        else if (click == "Right")
                                            cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.RightClick);
                                        else
                                            throw new System.Exception("Mouse Click Type wrong");
                                        break;
                                    }
                                }

                            }
                        }
                        if (cell.ChildNodes.ElementAt(0).InnerText != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).InnerText == Text)
                            {
                                foreach (HtmlTableCell cell2 in row.Cells)
                                {
                                    if (cell2.ChildNodes.ElementAt(0).InnerText == Text2)
                                    {
                                        RetunValue = true;
                                        HtmlTableCell cell1 = row.Cells[Column];
                                        if (scroll == true)
                                            cell1.ScrollToVisible();
                                        if (click == "Double")
                                            cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftDoubleClick);
                                        else if (click == "Single")
                                            cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                                        else if (click == "Right")
                                            cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.RightClick);
                                        else
                                            throw new System.Exception("Mouse Click Type wrong");
                                        break;
                                    }
                                }

                            }
                        }
                    }
                    else
                    {
                        if (cell.TextContent.ToString() == Text)
                        {
                            foreach (HtmlTableCell cell2 in row.Cells)
                            {
                                if (cell2.TextContent.ToString() == Text2)
                                {
                                    RetunValue = true;
                                    HtmlTableCell cell1 = row.Cells[Column];
                                    if (scroll == true)
                                        cell1.ScrollToVisible();
                                    if (click == "Double")
                                        cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftDoubleClick);
                                    else if (click == "Single")
                                        cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                                    else if (click == "Right")
                                        cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.RightClick);
                                    else
                                        throw new System.Exception("Mouse Click Type wrong");

                                    break;
                                }
                            }

                        }
                    }
                }
                if (RetunValue == true)
                {
                    break;
                }

            }
            return RetunValue;
        }

        /// <summary>
        /// This Function returns the cell from table from a specified column and the row which has the specified texts in same row..
        /// </summary>
        /// <param name="table">Pass the Htmltable object</param>
        /// <param name="text">Pass the string you are looking for in the table</param>
        /// <param name="text2">Pass the 2nd string you are looking for in the table</param>
        /// <param name="Column">Pass the column number of the table which has the button</param>
        /// <returns>The HTMLTable Cell required</returns>
        public static HtmlTableCell CellFromTableByMultiText(HtmlTable table, String Text, String Text2, int Column)
        {
            HtmlTableCell RetunValue = null;
            foreach (HtmlTableRow row in table.AllRows)
            {

                foreach (HtmlTableCell cell in row.Cells)
                {

                    if (cell.ChildNodes.Count >= 1)
                    {
                        if (cell.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).TextContent.ToString() == Text)
                            {
                                foreach (HtmlTableCell cell2 in row.Cells)
                                {
                                    if (cell2.ChildNodes.ElementAt(0).TextContent.ToString() == Text2)
                                    {

                                        HtmlTableCell cell1 = row.Cells[Column];
                                        RetunValue = cell1;

                                        break;
                                    }
                                }
                            }
                        }
                        if (cell.ChildNodes.ElementAt(0).InnerText != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).InnerText == Text)
                            {
                                foreach (HtmlTableCell cell2 in row.Cells)
                                {
                                    if (cell2.ChildNodes.ElementAt(0).InnerText == Text2)
                                    {

                                        HtmlTableCell cell1 = row.Cells[Column];
                                        RetunValue = cell1;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (cell.TextContent.ToString() == Text)
                        {
                            foreach (HtmlTableCell cell2 in row.Cells)
                            {
                                if (cell2.TextContent.ToString() == Text2)
                                {
                                    HtmlTableCell cell1 = row.Cells[Column];
                                    RetunValue = cell1;
                                    break;
                                }
                            }

                        }
                    }
                }
                if (RetunValue != null)
                {
                    break;
                }

            }
            return RetunValue;
        }

        /// <summary>
        /// This Function finds 2 values in a table in same row and return true if found.
        /// </summary>
        /// <param name="table">Pass the Htmltable object</param>
        /// <param name="text">Pass the string you are looking for in the table</param>
        /// <param name="text2">Pass the second string you are looking for in the table</param>
        /// <param name="Column">Pass the column number of the table which has the button</param>
        /// <returns>true if value found, else false</returns>
        public static Boolean FindMultiTextInTable(HtmlTable table, String Text, String Text2)
        {
            Boolean RetunValue = false;
            foreach (HtmlTableRow row in table.AllRows)
            {

                foreach (HtmlTableCell cell in row.Cells)
                {
                    if (cell.ChildNodes.Count >= 1)
                    {
                        if (cell.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).TextContent.ToString() == Text)
                            {
                                foreach (HtmlTableCell cell2 in row.Cells)
                                {
                                    if (cell2.ChildNodes.ElementAt(0).TextContent.ToString() == Text2)
                                    {
                                        RetunValue = true;
                                        break;
                                    }
                                }
                            }
                        }
                        if (cell.ChildNodes.ElementAt(0).InnerText != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).InnerText == Text)
                            {
                                foreach (HtmlTableCell cell2 in row.Cells)
                                {
                                    if (cell2.ChildNodes.ElementAt(0).InnerText == Text2)
                                    {
                                        RetunValue = true;
                                        break;
                                    }
                                }
                            }
                        }

                    }
                    else
                    {
                        if (cell.TextContent.ToString() == Text)
                        {
                            foreach (HtmlTableCell cell2 in row.Cells)
                            {
                                if (cell2.TextContent.ToString() == Text2)
                                {
                                    RetunValue = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                if (RetunValue == true)
                {
                    break;
                }

            }
            return RetunValue;
        }

        /// <summary>
        /// This Function Single Clicks on a button in a Html table according to a text entered if both text are in same row.
        /// </summary>
        /// <param name="table">Pass the Htmltable object</param>
        /// <param name="text">Pass the string you are looking for in the table</param>
        /// <param name="text2">Pass the second string you are looking for in the table</param>
        /// <param name="Column">Pass the column number of the table which has the button</param>
        /// <returns>Tru if value was found and clicked, else false</returns>
        public static Boolean ClickFromTableByMultiText(HtmlTable table, String Text,String Text2, int Column)
        {
            Boolean RetunValue = false;
            foreach (HtmlTableRow row in table.AllRows)
            {

                foreach (HtmlTableCell cell in row.Cells)
                {
                    if (cell.ChildNodes.Count >= 1)
                    {
                        if (cell.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).TextContent.ToString() == Text)
                            {
                                foreach (HtmlTableCell cell2 in row.Cells)
                                {
                                    if (cell2.ChildNodes.ElementAt(0).TextContent.ToString() == Text2)
                                    {
                                        RetunValue = true;
                                        HtmlTableCell cell1 = row.Cells[Column];

                                        cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                                        break;
                                    }
                                }

                            }
                        }
                        if (cell.ChildNodes.ElementAt(0).InnerText != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).InnerText == Text)
                            {
                                foreach (HtmlTableCell cell2 in row.Cells)
                                {
                                    if (cell2.ChildNodes.ElementAt(0).InnerText == Text2)
                                    {
                                        RetunValue = true;
                                        HtmlTableCell cell1 = row.Cells[Column];

                                        cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                                        break;
                                    }
                                }

                            }
                        }

                    }
                    else
                    {
                        if (cell.TextContent.ToString() == Text)
                        {
                            foreach (HtmlTableCell cell2 in row.Cells)
                            {
                                if (cell2.TextContent.ToString() == Text2)
                                {
                                    RetunValue = true;
                                    HtmlTableCell cell1 = row.Cells[Column];
                                    //cell1.ScrollToVisible();
                                    cell1.MouseClick(ArtOfTest.WebAii.Core.MouseClickType.LeftClick);
                                    break;
                                }
                            }


                        }
                    }
                }
                if (RetunValue == true)
                {
                    break;
                }

            }
            return RetunValue;
        }
        /// <summary>
        /// This Function returns the value on a column for the row selected by finding a 2 strings in the same row in the table.
        /// </summary>
        /// <param name="table">Pass the Htmltable object</param>
        /// <param name="text">Pass the string you are looking for in the table</param>
        /// <param name="text2">Pass the second string you are looking for in the table</param>
        /// <param name="Column">Pass the column number of the table which has the value we are looking for</param>
        /// <returns>Value on the required column</returns>
        public static String GetDataFromTableByMultiText(HtmlTable table, String Text,String Text2, int Column)
        {
            String RetunValue = "";
            Boolean flag = false;
            foreach (HtmlTableRow row in table.AllRows)
            {

                foreach (HtmlTableCell cell in row.Cells)
                {
                    if (cell.ChildNodes.Count >= 1)
                    {
                        if (cell.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).TextContent.ToString() == Text)
                            {
                        foreach (HtmlTableCell cell2 in row.Cells)
                            {
                                if (cell2.ChildNodes.ElementAt(0).TextContent.ToString() == Text2)
                                {
                                flag = true;
                                HtmlTableCell cell1 = row.Cells[Column];

                                if (cell1.ChildNodes.Count >= 1)
                                {
                                    if (cell1.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                                        RetunValue = cell1.ChildNodes.ElementAt(0).TextContent.ToString();
                                    else
                                        RetunValue = cell1.ChildNodes.ElementAt(0).InnerText.ToString();
                                }
                                else
                                {
                                    if (cell1.TextContent.ToString() != "")
                                        RetunValue = cell1.TextContent.ToString();
                                }


                                break;
                                }
                            }

                            }
                        }
                        if (cell.ChildNodes.ElementAt(0).InnerText != "")
                        {
                            if (cell.ChildNodes.ElementAt(0).InnerText == Text)
                            {
                            foreach (HtmlTableCell cell2 in row.Cells)
                            {
                                if (cell2.ChildNodes.ElementAt(0).InnerText == Text2)
                                {
                                flag = true;
                                HtmlTableCell cell1 = row.Cells[Column];

                                if (cell1.ChildNodes.Count >= 1)
                                {
                                    if (cell1.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                                        RetunValue = cell1.ChildNodes.ElementAt(0).TextContent.ToString();
                                    else
                                        RetunValue = cell1.ChildNodes.ElementAt(0).InnerText.ToString();
                                }
                                else
                                {
                                    if (cell1.TextContent.ToString() != "")
                                        RetunValue = cell1.TextContent.ToString();
                                }

                                break;
                                }
                                }

                            }
                        }

                    }
                    else
                    {
                        if (cell.TextContent.ToString() == Text)
                        {
                        foreach (HtmlTableCell cell2 in row.Cells)
                            {
                                if (cell2.TextContent.ToString()== Text2)
                                {
                            flag = true;
                            HtmlTableCell cell1 = row.Cells[Column];
                            //cell1.ScrollToVisible();
                            if (cell1.ChildNodes.Count >= 1)
                            {
                                if (cell1.ChildNodes.ElementAt(0).TextContent.ToString() != "")
                                    RetunValue = cell1.ChildNodes.ElementAt(0).TextContent.ToString();
                                else
                                    RetunValue = cell1.ChildNodes.ElementAt(0).InnerText.ToString();
                            }
                            else
                            {
                                if (cell1.TextContent.ToString() != "")
                                    RetunValue = cell1.TextContent.ToString();
                            }

                            break;
                                }
                            }


                        }
                    }
                }
                if (flag == true)
                {
                    break;
                }

            }
            return RetunValue;
        }
    }
}
    
