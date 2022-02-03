using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data;
using YahooFinanceApi;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btn_ShowPortofolio_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            activeSheet.Cells.ClearContents();

            SqlConnection con = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = C:\Irina\Facultate\Master\An I\Sem I\IPLA\Seminar\ExcelAddIn1-1113\bin\Debug\data.mdf; Integrated Security = True");
            try
            {
                if (con.State != ConnectionState.Open && con.State != ConnectionState.Connecting)
                {
                    con.Open();
                }

                SqlDataAdapter sda = new SqlDataAdapter("Select Stock, Amount from PORTOFOLIO", con);
                System.Data.DataTable dt = new System.Data.DataTable();

               

                sda.Fill(dt);
                if(dt.Rows.Count > 0)
                {
                    ((Excel.Range)activeSheet.Cells[1, 1]).Value2 = "Stock";
                    ((Excel.Range)activeSheet.Cells[1, 2]).Value2 = "Amount";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string stock = dt.Rows[i][0].ToString();
                        float amount = Convert.ToSingle(dt.Rows[i][1]);

                        Console.WriteLine(stock);
                        Console.WriteLine(amount);
                        ((Excel.Range)activeSheet.Cells[i+2, 1]).Value2 = stock;
                        ((Excel.Range)activeSheet.Cells[i+2, 2]).Value2 = amount.ToString();
                    }
                }

           

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {

                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
            }

        }

        private void btn_AddBuys_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            activeSheet.Cells.ClearContents();

            ((Excel.Range)activeSheet.Cells[1, 1]).Value2 = "Stock Exchange";
            ((Excel.Range)activeSheet.Cells[1, 2]).Value2 = "Stock";
            ((Excel.Range)activeSheet.Cells[1, 3]).Value2 = "Amount";
            ((Excel.Range)activeSheet.Cells[1, 4]).Value2 = "Price";
            ((Excel.Range)activeSheet.Cells[1, 5]).Value2 = "Date";

        }

        private void btn_SaveBuysDatabase_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            if(selection != null)
            {
                int row = selection.Row;
                int column = selection.Column;
                int row_count = selection.Rows.Rows.Count;
                int column_count = selection.Columns.Columns.Count;

                if (column_count == 4)
                {
                    string adress = (Globals.ThisAddIn.Application.Selection as Range).get_Address().ToString();
                    object[,] holder = selection.Value;
                    string[,] s = new string[row_count, column_count];
                    for (int q = 0; q < row_count; q++)
                    {

                        try
                        {
                            string stock = holder[q + 1, 1].ToString();
                            float amount = Convert.ToSingle(holder[q + 1, 2]);
                            float price = Convert.ToSingle(holder[q + 1, 3]);
                            DateTime date = (DateTime)holder[q + 1, 4];

                            SqlConnection con = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = C:\Irina\Facultate\Master\An I\Sem I\IPLA\Seminar\ExcelAddIn1-1113\bin\Debug\data.mdf; Integrated Security = True");
                            try
                            {
                                if (con.State != ConnectionState.Open && con.State != ConnectionState.Connecting)
                                {
                                    con.Open();
                                }
                                SqlCommand _cmd = new SqlCommand("INSERT INTO BUYS (STOCK,PRICE,DATE,AMOUNT) VALUES (@stock,@price,@date,@amount)", con);
                                _cmd.Parameters.AddWithValue("stock", stock);
                                _cmd.Parameters.AddWithValue("price", price);
                                _cmd.Parameters.AddWithValue("date", date);
                                _cmd.Parameters.AddWithValue("amount", amount);
                                _cmd.ExecuteNonQuery();
                            }
                            catch (Exception ex)
                            {
                               // System.Windows.Forms.MessageBox.Show(ex.Message);
                            }
                            finally
                            {

                                if (con.State == ConnectionState.Open)
                                {
                                    con.Close();
                                }

                               // System.Windows.Forms.MessageBox.Show("Buys Inserted into Database");

                            }

                        }
                        catch (Exception ex)
                        {
                           // System.Windows.Forms.MessageBox.Show(ex.Message);
                        }
                    }
                    System.Windows.Forms.MessageBox.Show("Buys Inserted into Database");
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Make sure you selected the right columns!");
                }
            }
           
        }

        private void btn_ShowAllBuys_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            activeSheet.Cells.ClearContents();

            SqlConnection con = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = C:\Irina\Facultate\Master\An I\Sem I\IPLA\Seminar\ExcelAddIn1-1113\bin\Debug\data.mdf; Integrated Security = True");
            try
            {
                if (con.State != ConnectionState.Open && con.State != ConnectionState.Connecting)
                {
                    con.Open();
                }

                SqlDataAdapter sda = new SqlDataAdapter("Select STOCK,PRICE, DATE, AMOUNT from BUYS", con);
                System.Data.DataTable dt = new System.Data.DataTable();



                sda.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    ((Excel.Range)activeSheet.Cells[1, 1]).Value2 = "Stock";
                    ((Excel.Range)activeSheet.Cells[1, 2]).Value2 = "Price";
                    ((Excel.Range)activeSheet.Cells[1, 3]).Value2 = "Amount";
                    ((Excel.Range)activeSheet.Cells[1, 4]).Value2 = "Date";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string stock = dt.Rows[i][0].ToString();
                        float price = Convert.ToSingle(dt.Rows[i][1]);
                        DateTime date = (DateTime) (dt.Rows[i][2]);
                        float amount = Convert.ToSingle(dt.Rows[i][3]);

                        ((Excel.Range)activeSheet.Cells[i + 2, 1]).Value = stock;
                        ((Excel.Range)activeSheet.Cells[i + 2, 2]).Value = price;
                        ((Excel.Range)activeSheet.Cells[i + 2, 3]).Value = amount;
                        ((Excel.Range)activeSheet.Cells[i + 2, 4]).NumberFormat = "mm/dd/yyyy";              
                        ((Excel.Range)activeSheet.Cells[i + 2, 4]).Value2 = date;                 
                    }
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {

                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
            }

        }

        private void btnShowStockEvolution_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;

            if (actCell.Value2 != null)
            {
                string symbol = actCell.Value2.ToString().ToUpper();
                string sText = actCell.Text;
        

                int timespan = 6;
                DateTime endDate = DateTime.Today;
                DateTime startDate = DateTime.Today.AddMonths(-timespan);
                StockData stock = new StockData(symbol);
                var awaiter = stock.getStockData( startDate, endDate);


                if (awaiter.Result == 1)
                {
                    Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
                    IReadOnlyList<Candle> historic_data = stock.Historic_Data;


                    if (historic_data != null && historic_data.Count > 0)
                    {
                        ((Excel.Range)activeSheet.Cells[1, 1]).Value2 = "Date";
                        ((Excel.Range)activeSheet.Cells[1, 2]).Value2 = "Close";
                        for (int i = 0; i < historic_data.Count; i++)
                        {
                            ((Excel.Range)activeSheet.Cells[i + 2, 1]).Value2 = historic_data[i].DateTime;
                            ((Excel.Range)activeSheet.Cells[i + 2, 1]).NumberFormat = "mm/dd/yyyy";
                            ((Excel.Range)activeSheet.Cells[i + 2, 2]).Value = historic_data[i].Close;

                        }

                        //Add Chart
                        Shape myChart = (Shape)sheet.Shapes.AddChart(XlChartType.xlLine, 275, 50, 400, 300);

                        var seriesOrig = (Series)myChart.Chart.SeriesCollection(1);
                        seriesOrig.Values = historic_data.Select(it => (double)it.Close).ToArray();
                        seriesOrig.XValues = historic_data.Select(it => it.DateTime).ToArray();
                        seriesOrig.Name = stock.Symbol;
                    }
                }
            }

        }

        private void btn_AddSales_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            activeSheet.Cells.ClearContents();

            ((Excel.Range)activeSheet.Cells[1, 1]).Value2 = "Stock Exchange";
            ((Excel.Range)activeSheet.Cells[1, 2]).Value2 = "Stock";
            ((Excel.Range)activeSheet.Cells[1, 3]).Value2 = "Amount";
            ((Excel.Range)activeSheet.Cells[1, 4]).Value2 = "Price";
            ((Excel.Range)activeSheet.Cells[1, 5]).Value2 = "Date";
        }

        private void btn_SaveSalesDatabase_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;

            if(selection != null)
            {
                int row = selection.Row;
                int column = selection.Column;
                int row_count = selection.Rows.Rows.Count;
                int column_count = selection.Columns.Columns.Count;

                if (column_count == 4)
                {
                    string adress = (Globals.ThisAddIn.Application.Selection as Range).get_Address().ToString();
                    object[,] holder = selection.Value;
                    string[,] s = new string[row_count, column_count];
                    for (int q = 0; q < row_count; q++)
                    {

                        try
                        {
                            string stock = holder[q + 1, 1].ToString();
                            float amount = Convert.ToSingle(holder[q + 1, 2]);
                            float price = Convert.ToSingle(holder[q + 1, 3]);
                            DateTime date = (DateTime)holder[q + 1, 4];

                            SqlConnection con = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = C:\Irina\Facultate\Master\An I\Sem I\IPLA\Seminar\ExcelAddIn1-1113\bin\Debug\data.mdf; Integrated Security = True");
                            try
                            {
                                if (con.State != ConnectionState.Open && con.State != ConnectionState.Connecting)
                                {
                                    con.Open();
                                }
                                SqlCommand _cmd = new SqlCommand("INSERT INTO SALES (STOCK,PRICE,DATE,AMOUNT) VALUES (@stock,@price,@date,@amount)", con);
                                _cmd.Parameters.AddWithValue("stock", stock);
                                _cmd.Parameters.AddWithValue("price", price);
                                _cmd.Parameters.AddWithValue("date", date);
                                _cmd.Parameters.AddWithValue("amount", amount);
                                _cmd.ExecuteNonQuery();
                            }
                            catch (Exception ex)
                            {
                               // System.Windows.Forms.MessageBox.Show(ex.Message);
                            }
                            finally
                            {

                                if (con.State == ConnectionState.Open)
                                {
                                    con.Close();
                                }

                               // System.Windows.Forms.MessageBox.Show("Sales Inserted into Database");

                            }

                        }
                        catch (Exception ex)
                        {

                        }
                    }
                    System.Windows.Forms.MessageBox.Show("Sales Inserted into Database");
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Make sure you selected the right columns!");
                }
            }
          
        }

        private void btn_ShowAllSales_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            activeSheet.Cells.ClearContents();

            SqlConnection con = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = C:\Irina\Facultate\Master\An I\Sem I\IPLA\Seminar\ExcelAddIn1-1113\bin\Debug\data.mdf; Integrated Security = True");
            try
            {
                if (con.State != ConnectionState.Open && con.State != ConnectionState.Connecting)
                {
                    con.Open();
                }

                SqlDataAdapter sda = new SqlDataAdapter("Select STOCK,PRICE, DATE, AMOUNT from SALES", con);
                System.Data.DataTable dt = new System.Data.DataTable();

                sda.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    ((Excel.Range)activeSheet.Cells[1, 1]).Value2 = "Stock";
                    ((Excel.Range)activeSheet.Cells[1, 2]).Value2 = "Price";
                    ((Excel.Range)activeSheet.Cells[1, 3]).Value2 = "Amount";
                    ((Excel.Range)activeSheet.Cells[1, 4]).Value2 = "Date";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string stock = dt.Rows[i][0].ToString();
                        float price = Convert.ToSingle(dt.Rows[i][1]);
                        DateTime date = (DateTime)(dt.Rows[i][2]);
                        float amount = Convert.ToSingle(dt.Rows[i][3]);

                        ((Excel.Range)activeSheet.Cells[i + 2, 1]).Value = stock;
                        ((Excel.Range)activeSheet.Cells[i + 2, 2]).Value = price;
                        ((Excel.Range)activeSheet.Cells[i + 2, 3]).Value = amount;
                        ((Excel.Range)activeSheet.Cells[i + 2, 4]).NumberFormat = "mm/dd/yyyy";
                        ((Excel.Range)activeSheet.Cells[i + 2, 4]).Value2 = date;
                    }
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {

                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
            }

        }

        private void btn_UpdatePortofolio_Click(object sender, RibbonControlEventArgs e)
        {
            SqlConnection con = new SqlConnection(@"Data Source = (LocalDB)\MSSQLLocalDB; AttachDbFilename = C:\Irina\Facultate\Master\An I\Sem I\IPLA\Seminar\ExcelAddIn1-1113\bin\Debug\data.mdf; Integrated Security = True");
            try
            {
                if (con.State != ConnectionState.Open && con.State != ConnectionState.Connecting)
                {
                    con.Open();
                }

                SqlDataAdapter sda = new SqlDataAdapter("Select STOCK, SUM(AMOUNT) from BUYS GROUP BY STOCK", con);
                System.Data.DataTable dt_BUYS = new System.Data.DataTable();

                Dictionary<string, float> portofolio = new Dictionary<string, float>();

                sda.Fill(dt_BUYS);
                if (dt_BUYS.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_BUYS.Rows.Count; i++)
                    {

                        portofolio.Add(dt_BUYS.Rows[i][0].ToString(), Convert.ToSingle(dt_BUYS.Rows[i][1]));
                    }
                }

                SqlDataAdapter sda1 = new SqlDataAdapter("Select STOCK, SUM(AMOUNT) from SALES GROUP BY STOCK", con);
                System.Data.DataTable dt_SALES = new System.Data.DataTable();

                sda1.Fill(dt_SALES);
                if (dt_SALES.Rows.Count > 0)
                {
                    for (int i = 0; i < dt_SALES.Rows.Count; i++)
                    {
                        if (portofolio.ContainsKey(dt_SALES.Rows[i][0].ToString()))
                        {
                            portofolio[dt_SALES.Rows[i][0].ToString()] = portofolio[dt_SALES.Rows[i][0].ToString()] - Convert.ToSingle(dt_SALES.Rows[i][1]);
                        }
                    }
                }
                Console.WriteLine(portofolio);


                for(int i = 0; i< portofolio.Count; i++)
                {

                   

                    SqlCommand cmd_checkIfStockIsInPortofolio = new SqlCommand("Select * from PORTOFOLIO WHERE Stock='"+portofolio.ElementAt(i).Key+"'", con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd_checkIfStockIsInPortofolio);
                    System.Data.DataTable ds = new System.Data.DataTable();
                    da.Fill(ds);
                    if( ds.Rows.Count > 0)
                    {
                        //update 
                        SqlCommand _cmd = new SqlCommand("UPDATE PORTOFOLIO SET Amount = @amount WHERE Stock = @stock", con);
                        _cmd.Parameters.AddWithValue("stock", portofolio.ElementAt(i).Key);
                        _cmd.Parameters.AddWithValue("amount", portofolio.ElementAt(i).Value);
                        _cmd.ExecuteNonQuery();
                    }
                    else
                    {
                        //insert
                        SqlCommand _cmd = new SqlCommand("INSERT INTO PORTOFOLIO (Stock,Amount) VALUES (@stock,@amount)", con);
                        _cmd.Parameters.AddWithValue("stock", portofolio.ElementAt(i).Key);
                        _cmd.Parameters.AddWithValue("amount", portofolio.ElementAt(i).Value);
                        _cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
            finally
            {

                if (con.State == ConnectionState.Open)
                {
                    con.Close();
                }
                System.Windows.Forms.MessageBox.Show("Portofolio updated succefully!");
            }
        }
    }
}
