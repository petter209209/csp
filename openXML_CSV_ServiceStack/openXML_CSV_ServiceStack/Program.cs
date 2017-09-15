using ServiceStack.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ServiceStack;
using System.Dynamic;

namespace openXML_CSV_ServiceStack
{
    class Program
    {
        
        static void Main(string[] args)
        {
            List<CSPInvoice> invoices=readDataCSV();            
            Console.ReadLine();

        }

        const int CUSTOMER_NAME_INDEX = 2;
        const int OFFWE_NAME_INDEX = 9;
        const int CHARGE_START_DATE_INDEX = 12;
        const int CHARGE_NED_DATE_INDEX = 13;
        const int CHARGE_TYPE_INDEX = 14;
        const int UNITP_RICE_INDEX = 15;
        const int QUANTITY_INDEX = 16;

        static List<CSPInvoice> readDataCSV()
        {
            //檔案位子
            string filePath = @"C:\Users\user\Desktop\testXML\ESi Technology_TWN_SRR_20170817-D030001OAF.CSV";
            var csv = File.ReadAllText(filePath);
            Console.WriteLine("載入Excel資料中...");
            string filename = FileName(filePath);
            //propNames標題列 rows資料內容
            string[] propNames = null;
            List<string[]> rows = new List<string[]>();

            //把資料裝進去rows
            foreach (var line in CsvReader.ParseLines(csv))
            {
                string[] strArray = CsvReader.ParseFields(line).ToArray();
                if (propNames == null)
                    propNames = strArray;
                else
                    rows.Add(strArray);
            }
            //Console.WriteLine($"PropNames={string.Join(",", propNames)}");
            Console.WriteLine($"載入Excel成功，共{rows.Count}筆");

            //要輸出的資料
            List<CSPInvoice> invoices = GetCSPInvoiceFromCSV(rows);

            //show data
            //showdata(invoices);

            Console.WriteLine("資料上傳中");
            // ImportDynamics
            ImportDynamics importDynamics = new ImportDynamics(invoices, filename);
            importDynamics.startImportDynamics();

            Console.WriteLine("資料上傳完畢");
            return invoices;
        }
        static int differenceDate(DateTime oldDate,DateTime newDate)
        {
            //計算差距天數

            // Difference in days, hours, and minutes.
            TimeSpan ts = newDate - oldDate;
            // Difference in days.
            int differenceInDays = ts.Days;
            return differenceInDays;
        }
        static int monthDays(DateTime targetmonth)
        {
            //計算當月共有幾天
            DateTime FirstDay = targetmonth.AddDays(-DateTime.Now.Day + 1);
            DateTime LastDay = targetmonth.AddMonths(1).AddDays(-DateTime.Now.AddMonths(1).Day);
            return differenceDate(FirstDay, LastDay);
        }
        static List<CSPInvoice> GetCSPInvoiceFromCSV(List<String[]> rows)
        {
            //建立 SCPInvoice
            List<CSPInvoice> invoices = new List<CSPInvoice>();

            //把資料依序塞選進入 List<CSPInvoice>
            for (int r = 0; r < rows.Count; r++)
            {
                var cells = rows[r];
                //日期相關資
                //initial 初始資料
 
                DateTime startDate = Convert.ToDateTime(cells[CHARGE_START_DATE_INDEX]); ;
                DateTime endDate = Convert.ToDateTime(cells[CHARGE_NED_DATE_INDEX]); ;

                //建立新 CSPInvoice
                CSPInvoice invoice = new CSPInvoice();
                //塞入資料
                invoice.CustomerName = cells[CUSTOMER_NAME_INDEX];
                invoice.OfferName = cells[OFFWE_NAME_INDEX];
                invoice.ChargeStartDate = startDate;
                invoice.ChargeEndDate = endDate;
                invoice.ChargeType = cells[CHARGE_TYPE_INDEX];
                invoice.UnitPrice = cells[UNITP_RICE_INDEX];
                invoice.Quantity = cells[QUANTITY_INDEX];
                invoice.ChargeDaterange = differenceDate(startDate, endDate);
                invoice.ChargeStartMonthDays = monthDays(startDate);


                //把 CSPInvoice 塞入 List<CSPInvoice>
                invoices.Add(invoice);
            }
            return invoices;
        }
        static void showdata(List<CSPInvoice> invoices)
        {
            //show data
                foreach (var p in invoices)
                {
                    Console.WriteLine("---------------------------------------------");
                    Console.WriteLine($"{"CustomerName",-20} : {p.CustomerName,-20}");
                    Console.WriteLine($"{"OfferName",-20} : {p.OfferName,-20}");
                    Console.WriteLine($"{"ChargeStartDate",-20} : {p.ChargeStartDate.ToString("yyyy/MM/dd"),-20}");
                    Console.WriteLine($"{"ChargeEndDate",-20} : {p.ChargeEndDate.ToString("yyyy/MM/dd"),-20}");
                    Console.WriteLine($"{"ChargeType",-20} : {p.ChargeType,-20}");
                    Console.WriteLine($"{"UnitPrice",-20} : {p.UnitPrice,-20}");
                    Console.WriteLine($"{"Quantity",-20} : {p.Quantity,-20}");
                    Console.WriteLine($"{"ChargeDaterange",-20} : {p.ChargeDaterange,-20}");
                    Console.WriteLine($"{"ChargeStartMonthDays",-20} : {p.ChargeStartMonthDays,-20}");
                }
           

        }
        static string FileName(string filename)
        {
            char[] delimiterChars = {'.'};
            filename = Path.GetFileName(filename);
            string[] filenames = filename.Split(delimiterChars);
            filename = filenames[0];
            return filename;
        }
    }
}
