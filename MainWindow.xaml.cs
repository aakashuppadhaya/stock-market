using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;
using LiveCharts;
using LiveCharts.Defaults;
using LiveCharts.Wpf;
using System.Windows.Annotations;
using System.ComponentModel;
using System.Runtime.InteropServices;
using excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data;
using Xceed.Wpf.Toolkit;



namespace Wpf.CartesianChart.Financial
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public class NameValidator : ValidationRule
    {
        public override ValidationResult Validate(object value, System.Globalization.CultureInfo cultureInfo)
        {
            if (value == null)
                return new ValidationResult(false, "value cannot be empty.");
            else
            {
                if (value.ToString().Length > 3)
                    return new ValidationResult(false, "Name cannot be more than 3 characters long.");
            }
            return ValidationResult.ValidResult;
        }
    }
    public partial class OhclExample : UserControl 
    {
        private string[] _labels;
        OleDbConnection oledbConn;
        decimal[] result1 = new decimal[400];
        double[] result2 = new double[400];
        string[] profit = new string[1000];
        string[] trades = new string[1000];
        string[] status = new string[1000];
        decimal[] Open = new decimal[1000];
        decimal[] High = new decimal[1000];
        decimal[] Close = new decimal[1000];
        decimal[] Low = new decimal[1000];
        DateTime[] Date = new DateTime[1000];
        float Buy = 0, Sell = 0, Hold = 0;
        public OhclExample()
        {
            InitializeComponent();
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            string path = "C:/project/Companies.xls";

            if (Path.GetExtension(path) == ".xls")
            {
                oledbConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"");
            }

            oledbConn.Open();
            OleDbCommand cmd = new OleDbCommand(); ;
            OleDbDataAdapter oleda = new OleDbDataAdapter();
            DataSet ds = new DataSet();

            cmd.Connection = oledbConn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT  * FROM [Hdfc$]";
            oleda = new OleDbDataAdapter(cmd);
            oleda.Fill(ds);
            DataTable dt = ds.Tables[0];

            decimal[] terms = new decimal[400];
            decimal[] shares = new decimal[400];
            decimal[] trade = new decimal[400];
            for (int i = 1; i < 400; i++)
            {
                terms[i] = Convert.ToDecimal(dt.Rows[i]["Total Turnover"]);
                shares[i] = Convert.ToDecimal(dt.Rows[i]["No# of Shares"]);
                trade[i] = Convert.ToDecimal(dt.Rows[i]["No# of Trades"]);

            }

            
             
           
          
         
        
        }

       
        private void UpdateAllOnClick(object sender, RoutedEventArgs e)
        {

            CreateCompanyData();
           
            //OHLC GRAPH CODE

            SeriesCollection = new SeriesCollection()
            {
                new OhlcSeries()
                {
  
                    Values = new ChartValues<OhlcPoint>
                    {
                              new OhlcPoint ((double)Open[0],(double)High[0],(double)Close[0],(double)Low[0]),
                             new OhlcPoint ((double)Open[1],(double)High[1],(double)Close[1],(double)Low[1]),
                             new OhlcPoint ((double)Open[2],(double)High[2],(double)Close[2],(double)Low[2]),
                             new OhlcPoint ((double)Open[3],(double)High[3],(double)Close[3],(double)Low[3]),
                             new OhlcPoint ((double)Open[4],(double)High[4],(double)Close[4],(double)Low[4]),
                             new OhlcPoint ((double)Open[5],(double)High[5],(double)Close[5],(double)Low[5]),
                             new OhlcPoint ((double)Open[6],(double)High[6],(double)Close[6],(double)Low[6]),
                             new OhlcPoint ((double)Open[7],(double)High[7],(double)Close[7],(double)Low[7]),
                             new OhlcPoint ((double)Open[8],(double)High[8],(double)Close[8],(double)Low[8]),
                             new OhlcPoint ((double)Open[9],(double)High[9],(double)Close[9],(double)Low[9]),
                             new OhlcPoint ((double)Open[10],(double)High[10],(double)Close[10],(double)Low[10]),
                             new OhlcPoint ((double)Open[11],(double)High[11],(double)Close[11],(double)Low[11]),
                             new OhlcPoint ((double)Open[12],(double)High[12],(double)Close[12],(double)Low[12]),
                             new OhlcPoint ((double)Open[13],(double)High[13],(double)Close[13],(double)Low[13]),
                             new OhlcPoint ((double)Open[14],(double)High[14],(double)Close[14],(double)Low[14]),
                       /* new OhlcPoint(1032, 1035, 1030, 1032),
                        new OhlcPoint(1033, 1038, 1031, 1037),
                        new OhlcPoint(1035, 1042, 1030, 1040),
                        new OhlcPoint(1037, 1040, 1035, 1038),
                        new OhlcPoint(1035, 1038, 1032, 1033),
                        new OhlcPoint(1042, 1055, 1050, 1042),
                        new OhlcPoint(1053, 1048, 1045, 1050),
                        new OhlcPoint(1055, 1052, 1040, 1050),
                        new OhlcPoint(1047, 1040, 1035, 1048),
                        new OhlcPoint(1035, 1044, 1032, 1033),
                        new OhlcPoint(1005, 1041, 1020, 1038)*/
                    }
                },
                
            };
           
            Labels = new[]
            {
       
                
                    Date[0].ToString("dd MM "),
                    Date[1].ToString("dd MM "),
                    Date[2].ToString("dd MM "),
                    Date[3].ToString("dd MM "),
                    Date[4].ToString("dd MM "),
                   
                    Date[5].ToString("dd MM "),
                    Date[6].ToString("dd MM "),
                    Date[7].ToString("dd MM "),
                    Date[8].ToString("dd MM "),
                    Date[9].ToString("dd MM "),
                    Date[10].ToString("dd MM"),
                    Date[11].ToString("dd MM"),
                    Date[12].ToString("dd MM"),
                    Date[13].ToString("dd MM"),
                    Date[14].ToString("dd MM"),
                    

               

                 /*DateTime.Now.AddDays(-11).ToString("dd MMM"),
                 DateTime.Now.AddDays(-10).ToString("dd MMM"),
                 DateTime.Now.AddDays(-9).ToString("dd MMM"),
                 DateTime.Now.AddDays(-8).ToString("dd MMM"),
                 DateTime.Now.AddDays(-7).ToString("dd MMM"),
                 DateTime.Now.AddDays(-6).ToString("dd MMM"),
                 DateTime.Now.AddDays(-5).ToString("dd MMM"),
                 DateTime.Now.AddDays(-4).ToString("dd MMM"),
                 DateTime.Now.AddDays(-3).ToString("dd MMM"),
                 DateTime.Now.AddDays(-2).ToString("dd MMM"),
                 DateTime.Now.AddDays(0).ToString("dd MMM"),
                 DateTime.Now.ToString("dd MMM"),
                 DateTime.Now.AddDays(1).ToString("dd MMM"),*/
              

                         
            };

            DataContext = this;
 
            var r = new Random();

            foreach (var point in SeriesCollection[0].Values.Cast<OhlcPoint>())
            {
                point.Open = r.Next((int)point.Low, (int)point.High);
                point.Close = r.Next((int)point.Low, (int)point.High);
                
            }

         //   clearChart(point.Open,point.Close);

        }

       /* private void clearChart(Values)
        {

            foreach (var point in SeriesCollection[0].Values.Cast<OhlcPoint>())
            { 
                Values.Clear();
            }
        }*/
        public SeriesCollection SeriesCollection { get; set; }

        public string[] Labels
        {
            get { return _labels; }
            set
            {
                _labels = value;
                OnPropertyChanged("Labels");
            }
        }

       
       public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            if (PropertyChanged != null) PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CreateCompanyData();
        }
        private void CreateCompanyData()
        {

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            string path = "C:/project/Companies.xls";

            if (Path.GetExtension(path) == ".xls")
            {
                oledbConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"");
            }

            oledbConn.Open();
            OleDbCommand cmd = new OleDbCommand(); ;
            OleDbDataAdapter oleda = new OleDbDataAdapter();
            DataSet ds = new DataSet();

            cmd.Connection = oledbConn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT  * FROM [" + comboBox1.SelectedValue + "$]";
            oleda = new OleDbDataAdapter(cmd);
            oleda.Fill(ds);
            DataTable dt = ds.Tables[0];

            decimal[] terms = new decimal[1000];
            decimal[] shares = new decimal[1000];
            decimal[] trade = new decimal[1000];
            for (int i = 0; i < 1000; i++)
            {
                terms[i] = Convert.ToDecimal(dt.Rows[i]["Total Turnover"]);
                shares[i] = Convert.ToDecimal(dt.Rows[i]["No# of Shares"]);
                trade[i] = Convert.ToDecimal(dt.Rows[i]["No# of Trades"]);
                Open[i] = Convert.ToDecimal(dt.Rows[i]["Open"]);
               High[i] = Convert.ToDecimal(dt.Rows[i]["High"]);
                Close[i] = Convert.ToDecimal(dt.Rows[i]["Close"]);
                Low[i] = Convert.ToDecimal(dt.Rows[i]["Low"]);
                Date[i] = Convert.ToDateTime(dt.Rows[i]["Date"]);

            }


            result1 = Get_N_DaysMovingAverage(50, terms);


            result2 = Get_N_DaysTrading(50, shares, trade);

            xyz(result1, result2);
            // abc(result2);
            //BuyOrSell(result1, result2);

        }

        private void xyz(decimal[] result1, double[] result2)
        {
            decimal turnoverSum = 0;
            double TradingSum = 0;
            double TradingAvg = 0;
            decimal turnoverAvg = 0;
            for (int i = 0; i < result1.Length; i++)
            {
                turnoverSum = turnoverSum + result1[i];
            }

            turnoverAvg = turnoverSum / result1.Length;
            for (int i = 0; i < result2.Length; i++)
            {
                TradingSum = TradingSum + result2[i];
            }

            TradingAvg = TradingSum / result2.Length;

            for (int i = 0; i < result1.Length; i++)
            {
                if (result1[i] < turnoverAvg && result2[i] < TradingAvg)
                {
                    profit[i] = "Low";
                    trades[i] = "Low";
                    status[i] = "Sell";
                }
                else if (result1[i] < turnoverAvg && result2[i] > TradingAvg)
                {
                    profit[i] = "Low";
                    trades[i] = "High";
                    status[i] = "Buy";
                }
                else if (result1[i] > turnoverAvg && result2[i] < TradingAvg)
                {
                    profit[i] = "High";
                    trades[i] = "Low";
                    status[i] = "Buy";
                }
                else if (result1[i] > turnoverAvg && result2[i] > TradingAvg)
                {
                    profit[i] = "High";
                    trades[i] = "High";
                    status[i] = "Sell";
                }

            }

        }

        /* private void abc(double[] result2)
         {
            double TradingSum = 0;
          
          
            double TradingAvg = 0;
             for (int i = 0; i < result2.Length; i++)
             {
                TradingSum = TradingSum + result2[i];
             }

             TradingAvg = TradingSum / result2.Length;

             for (int i = 0; i < result2.Length; i++)
             {
                 if (result2[i] < TradingAvg)
                 {
                     trades[i] = "Low";
                 }
                 else
                 {
                     trades[i] = "High";
                 }
             }



         }*/

        /*private void BuyOrSell(decimal[] result1, double[] result2)
        { 
           
        }*/


        private static decimal[] Get_N_DaysMovingAverage(int frameSize, decimal[] data)
        {//Moving average for analysis

            decimal sum = 0;
            decimal[] avgPoints = new decimal[data.Length - frameSize + 1];
            for (int counter = 0; counter <= data.Length - frameSize; counter++)
            {
                int innerLoopCounter = 0;
                int index = counter;
                while (innerLoopCounter < frameSize)
                {
                    sum = sum + data[index];

                    innerLoopCounter += 1;

                    index += 1;

                }

                avgPoints[counter] = sum / frameSize;

                sum = 0;

            }
            return avgPoints;

        }


        private static double[] Get_N_DaysTrading(int size, decimal[] share, decimal[] trading)
        {
            double[] trades = new double[trading.Length - size + 1];
            double qtyTraded = 0;
            for (int counter = 0; counter <= trading.Length - size; counter++)
            {
                int innerLoopCounter = 0;
                int index = counter;
                while (innerLoopCounter < size)
                {
                    double abbcqtyTraded = (double)trading[index] / (double)share[index];
                    qtyTraded = qtyTraded + abbcqtyTraded;

                    innerLoopCounter += 1;

                    index += 1;

                }
                trades[counter] = qtyTraded / size;

                qtyTraded = 0;

            }

            return trades;
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {

            CreateCompanyData();
            xyz(result1, result2);
           

            try
            {
                string filename = @"C:\ExcelLibrary\" + textBox1.Text + ".xls";
                if (File.Exists(filename))
                {

                }
                else
                {
                    Microsoft.Office.Interop.Excel._Application oApp;
                    Microsoft.Office.Interop.Excel._Worksheet oSheet;
                    Microsoft.Office.Interop.Excel._Workbook oBook;

                    oApp = new Microsoft.Office.Interop.Excel.Application();
                    oBook = oApp.Workbooks.Add();
                    oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oBook.Worksheets.get_Item(1);


                    oSheet.Cells[1, 1] = "TurnoverAvg";
                    oSheet.Cells[1, 2] = "TradingAvg";
                    oSheet.Cells[1, 3] = "Turnover";
                    oSheet.Cells[1, 4] = "Trading";
                    oSheet.Cells[1, 5] = "Status";

                    if (oApp.Application.Sheets.Count < 1)
                    {
                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oBook.Worksheets.Add();
                    }
                    else
                    {
                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oApp.Worksheets[1];
                    }



                    oBook.SaveAs(filename);
                    oBook.Close();
                    oApp.Quit();

                }
            }
            catch
            {

            }


          

            try
            {

                System.Data.OleDb.OleDbConnection MyConnection;
                System.Data.OleDb.OleDbCommand myCommand = new System.Data.OleDb.OleDbCommand();

                string path = "C:/ExcelLibrary/" + textBox1.Text + ".xls";
                MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;");
                MyConnection.Open();
                myCommand.Connection = MyConnection;
                for (int i = 0; i < result1.Length; i++)
                {
                    string myval1 = result1[i].ToString();
                    string myval2 = result2[i].ToString();
                    string myval3 = profit[i];
                    string myval4 = trades[i];
                    string myval5 = status[i];

                    myCommand = new System.Data.OleDb.OleDbCommand("Insert into [Sheet1$] (TurnoverAvg,TradingAvg,Turnover,Trading,Status) values('" + myval1 + "','" + myval2 + "','" + myval3 + "','" + myval4 + "','" + myval5 + "')", MyConnection);
                    // myCommand.CommandText = sql;
                    myCommand.ExecuteNonQuery();

                }

                MyConnection.Close();

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }


      private void button2_Click(object sender, EventArgs e)
      {
          //Prediction code using navie bayesian

          float THS = 0, THB = 0, THH = 0, TLS = 0, TLB = 0, TLH = 0, TradeHB = 0, TradeHS = 0, TradeHH=0, TradeLS = 0, TradeLB = 0, TradeLH = 0; //ay=age yes,an=age no;
          try
          {
           Excel.Application xlApp;
          Excel.Workbook xlWorkBook;
          Excel.Worksheet xlWorkSheet;
          Excel.Range range;
          string path = "C:/ExcelLibrary/" + textBox1.Text + ".xls";

          if (Path.GetExtension(path) == ".xls")
          {
              oledbConn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"");

          }
          oledbConn.Open();
          OleDbCommand cmd = new OleDbCommand(); ;
          OleDbDataAdapter oleda = new OleDbDataAdapter();
          DataSet ds = new DataSet();

          cmd.Connection = oledbConn;
          cmd.CommandType = CommandType.Text;
          cmd.CommandText = "SELECT  * FROM [Sheet1$]";
          oleda = new OleDbDataAdapter(cmd);
          oleda.Fill(ds);
          DataTable dt = ds.Tables[0];


          cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where Status='Buy';", oledbConn);
          cmd.ExecuteNonQuery();
          OleDbDataReader dr;
          dr = cmd.ExecuteReader();
          while (dr.Read())
          {
              Buy++;
          }
          dr.Close();

          cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where Status='Sell';", oledbConn);
          cmd.ExecuteNonQuery();
          dr = cmd.ExecuteReader();
          while (dr.Read())
          {
              Sell++;
          }
          dr.Close();

        /*  cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where Status='Hold';", oledbConn);
          cmd.ExecuteNonQuery();
          dr = cmd.ExecuteReader();
          while (dr.Read())
          {
              Hold++;
          }
          dr.Close();*/

          cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Turnover]='High' and [Status]='Sell';", oledbConn);
              cmd.ExecuteNonQuery();
              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  THS++; //calculation of total number yes with age choosen
              }
              dr.Close();
              cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Turnover]='High' And [Status]='Buy';", oledbConn);
              cmd.ExecuteNonQuery();
              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  THB++;
              }
              dr.Close();
            /*  cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Turnover]='High' And [Status]='Hold';", oledbConn);
              cmd.ExecuteNonQuery();
              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  THH++;
              }
              dr.Close();*/
              cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Turnover]='Low' And Status='Sell';", oledbConn);
              cmd.ExecuteNonQuery();

              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  TLS++;
              }
             dr.Close();
             cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Turnover]='Low' And Status='Buy';", oledbConn);
              cmd.ExecuteNonQuery();

              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  TLB++; 
              }
              dr.Close();
             /* cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Turnover]='Low' And [Status]='Hold';", oledbConn);
              cmd.ExecuteNonQuery();
              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  TLH++;
              }
              dr.Close();*/


              cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Trading]='High' and [Status]='Sell';", oledbConn);
              cmd.ExecuteNonQuery();
              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  TradeHS++; //calculation of total number yes with age choosen
              }
              dr.Close();
              cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Trading]='High' And [Status]='Buy';", oledbConn);
              cmd.ExecuteNonQuery();
              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  TradeHB++;
              }
              dr.Close();
              /*cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Trading]='High' And [Status]='Hold';", oledbConn);
              cmd.ExecuteNonQuery();
              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  TradeHH++;
              }
              dr.Close();*/
              cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Trading]='Low' And Status='Sell';", oledbConn);
              cmd.ExecuteNonQuery();

              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  TradeLS++;
              }
              dr.Close();
              cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Trading]='Low' And Status='Buy';", oledbConn);
              cmd.ExecuteNonQuery();

              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  TradeLB++;
              }
              dr.Close();
              /*cmd = new System.Data.OleDb.OleDbCommand("Select * from [Sheet1$] where [Trading]='Low' And [Status]='Hold';", oledbConn);
              cmd.ExecuteNonQuery();
              dr = cmd.ExecuteReader();
              while (dr.Read())
              {
                  TradeLH++;
              }
              dr.Close();*/


              float als = THS / Sell;   //aly=age learning phase of yes
              float alb = THB /Buy;
            //  float alh = THH / Hold;
              float ils = TLS / Sell;
              float ilb = TLB / Buy;
             // float ilh = TLH / Hold;
              float sls = TradeHS / Sell;
              float slb = TradeHB / Buy;
             // float slh = TradeHH / Hold;
              float cls = TradeLS / Sell;
              float clb = TradeLB / Buy;
             // float clh = TradeLH / Hold;

              float total = Buy + Sell;
              float ps = Sell / total;  //py=probability of yes
              float pb = Buy / total;
             // float ph = Hold / total;
              float pxs = als * ils * ps * sls * cls ; //pxy=probability of selling shares
              float pxb = alb * ilb * pb * slb * clb;
              //float pxh = alh* ilh*slh*clh;
              if (pxs < pxb)
              {
                  textBox2.Text = "You can 'BUY' the shares of this company";
              }
              else
              {
                  textBox2.Text = "You can 'SELL' the shares of this company";
              }
          }
          catch (Exception ex)
          {
             // MessageBox.Show(ex.Message);
          }
        }


        private void button3_Click(object sender, RoutedEventArgs e)
        {
            textBox1.Text = "Enter the file name";
            textBox2.Text = "";
            comboBox1.SelectedValue = "Please select the company";
            
        }

        private void textBox3_TextChanged(object sender, TextChangedEventArgs e)
        {

        }


       
       
    
       
    }
}
