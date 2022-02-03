using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YahooFinanceApi;
using System.Linq;
using System.Collections;

namespace ExcelAddIn1
{
    class StockData
    {
        private string _symbol;
        private IReadOnlyList<Candle> _historic_Data;

        public StockData(string symbol)
        {
            Symbol = symbol; 
        }

        public string Symbol { get => _symbol; set => _symbol = value; }
        public IReadOnlyList<Candle> Historic_Data { get => _historic_Data;}

        public async Task<int> getStockData(DateTime startDate, DateTime endDate)
        {
            try
            {
                var historic_data= await Yahoo.GetHistoricalAsync(Symbol, startDate, endDate);
                var security = await Yahoo.Symbols(Symbol).Fields(Field.LongName).QueryAsync();
                var ticker = security[Symbol];
                var companyName = ticker[Field.LongName];

                _historic_Data = new List<Candle>();
                _historic_Data =  historic_data;

                for (int i=0; i< historic_data.Count; i++)
                {
                    Console.WriteLine(companyName + "Closing price on:" + historic_data.ElementAt(i).DateTime.Month + "/" + historic_data.ElementAt(i).DateTime.Day + "/" + historic_data.ElementAt(i).DateTime.Year + ": $" + Math.Round(historic_data.ElementAt(i).Close,2));
           
                }
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Failed to get symbol: " + Symbol);
            }
            return 1;
        }
    }
}
