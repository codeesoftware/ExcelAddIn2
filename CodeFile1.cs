
using ExcelAddIn2;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using www.mnb.hu.webservices;
using Excel = Microsoft.Office.Interop.Excel;
public interface IFileOperations
{
    bool Save();
    bool Open();
    bool Read();
}
public class ExcelOperations : IFileOperations
{
    //    public string path;
    //    public string filename;
    public string fileNameWithPath;
    public System.Data.DataTable dataTable;

    public ExcelOperations(string _fileNameWithPath)
    {
        fileNameWithPath = _fileNameWithPath;
    }


    public bool Save()
    {
        try
        {
            GetCurrenciesResponseBody currenciesResponse = null;
            GetCurrencyUnitsResponseBody currenciesUnitResponse = null;
            GetExchangeRatesResponseBody bodycurrentExchangeRates = null;
            MNBCurrencies currencies = null;
            using (var client = new MNBArfolyamServiceSoapClient("CustomBinding_MNBArfolyamServiceSoap"))
            {
                var currencyRequestParameter = new GetCurrenciesRequestBody();
                currenciesResponse = client.GetCurrencies(currencyRequestParameter);

                XmlSerializer ser1 = new XmlSerializer(typeof(MNBCurrencies));
                using (var sr = new StringReader(currenciesResponse.GetCurrenciesResult))
                using (XmlReader reader = XmlReader.Create(sr))
                {
                    currencies = (MNBCurrencies)ser1.Deserialize(reader);
                }
                string currenciesString = String.Join(",", currencies.Currencies);

                var currencyUnitRequestParameter = new GetCurrencyUnitsRequestBody() { currencyNames = currenciesString };
                currenciesUnitResponse = client.GetCurrencyUnits(currencyUnitRequestParameter);


                var bodycurrentExchangeRatesRequestParameter = new GetExchangeRatesRequestBody()
                {
                    currencyNames = currenciesString,
                    startDate = "2017.10.02",
                    endDate = "2017.11.16.",
                };
                bodycurrentExchangeRates = client.GetExchangeRates(bodycurrentExchangeRatesRequestParameter);
            }
            XmlSerializer ser2 = new XmlSerializer(typeof(MNBCurrencyUnits));
            MNBCurrencyUnits currencyUnits;
            using (var sr = new StringReader(currenciesUnitResponse.GetCurrencyUnitsResult))
            using (XmlReader reader = XmlReader.Create(sr))
            {
                currencyUnits = (MNBCurrencyUnits)ser2.Deserialize(reader);
            }

            XmlSerializer ser3 = new XmlSerializer(typeof(MNBExchangeRates));
            MNBExchangeRates exchangeRates;
            using (var sr = new StringReader(bodycurrentExchangeRates.GetExchangeRatesResult))
            using (XmlReader reader = XmlReader.Create(sr))
            {
                exchangeRates = (MNBExchangeRates)ser3.Deserialize(reader);
            }

            DataTable table1 = new DataTable("Arfolyamok");
            table1.Columns.Add("Dátum/ISO");
            foreach (string currency in currencies.Currencies)
            {
                table1.Columns.Add(currency);

            }
            var unitDic = new Dictionary<string, int>();
            DataRow unitRow = table1.NewRow();
            unitRow[0] = "Egység";
            for (int i = 1; i < currencyUnits.Units.Length; i++)
            {
                var unit = currencyUnits.Units[i];
                unitRow[i] = unit.Value;
                unitDic.Add(unit.curr, i + 1);
            }
            table1.Rows.Add(unitRow);


            for (int i = exchangeRates.Day.Length -1 ; 0 < i ; i--)
            {


                DataRow exchangeRateDayRow = table1.NewRow();
                MNBExchangeRatesDay day = exchangeRates.Day[i];
                exchangeRateDayRow[0] = day.date;
                for (int j = 1; j < day.Rate.Length; j++)
                {
                    int columnIndex = unitDic[day.Rate[j].curr];
                    exchangeRateDayRow[columnIndex] = day.Rate[j].Value;
                }
                table1.Rows.Add(exchangeRateDayRow);

            }
            DataSet set = new DataSet("office");
            set.Tables.Add(table1);

            Export(set);
        }
        catch (System.Exception e)
        {

            throw;
        }
        return true;
    }
    public bool Open()
    {
        var _excel = new Excel.Application();
        var _workBook = _excel.Workbooks.Open(fileNameWithPath);
        return true;
    }
    public bool Read()
    {
        return true;
    }

    void Export(DataSet ds)
    {
        Excel.Application excelApp = new Excel.Application();

        Excel.Workbook excelWorkBook = excelApp.Workbooks.Add();

        foreach (DataTable table in ds.Tables)
        {
            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
            excelWorkSheet.Name = table.TableName;

            for (int i = 1; i < table.Columns.Count + 1; i++)
            {
                excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
            }

            for (int j = 0; j < table.Rows.Count; j++)
            {
                for (int k = 0; k < table.Columns.Count; k++)
                {
                    excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                }
            }
        }

        excelWorkBook.SaveAs(@"Sanyi.xlsx");
        excelWorkBook.Close();
        excelApp.Quit();

    }
}
public class AccessActions
{
    private string fileNameWithPath;
    private readonly Excel.Application appliction;

    public AccessActions(string _fileNameWithPath)
    {
        fileNameWithPath = _fileNameWithPath;
    }


    public bool ReadDataFromFile()
    {

        OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileNameWithPath);
        connection.Open();
        OleDbDataReader reader = null;
        OleDbCommand command = new OleDbCommand("SELECT * from Table1", connection);
        reader = command.ExecuteReader();
        if (!reader.HasRows)
        {
            connection.Close();
            return false;
        }

        var _workBook = ExcelAddIn2.Globals.ThisAddIn.Application.ActiveWorkbook;
        var _sheet = _workBook.Worksheets.Add(_workBook.Sheets[_workBook.Sheets.Count], 1);
        var _cells = _sheet.Cells;
        _sheet.Name = "Log";

        int rowCounter = 1;
        while (reader.Read())
        {
            int colCounter = 1;
            foreach (var field in reader)
            {
                _cells[rowCounter, colCounter].Value = field;
            }

        }

        connection.Close();
        return true;
    }
}
