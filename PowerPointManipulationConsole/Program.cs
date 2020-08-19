using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Powerpoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointManipulationConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var bankAccounts = new List<Account>
            {
                new Account
                {
                    ID = 15014475,
                    Balance = 450.300
                },
                new Account
                {
                    ID = 15014476,
                    Balance = -125.400
                }
            };

            DisplayInExcel(bankAccounts);

            WordDoc();
        }

        static void DisplayInExcel(IEnumerable<Account> accounts)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            workSheet.Cells[1, "A"] = "ID number";
            workSheet.Cells[1, "B"] = "Current Balance";

            var row = 1;

            foreach (var acc in accounts)
            {
                row++;
                workSheet.Cells[row, "A"] = acc.ID;
                workSheet.Cells[row, "B"] = acc.Balance;
            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();

            workSheet.Range["A1:B3"].AutoFormat ( Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2 );

            workSheet.Range["A1:B3"].Copy();

        }

        static void WordDoc()
        {
            var wordApp = new Word.Application();
            wordApp.Visible = true;
            wordApp.Documents.Add();
            wordApp.Selection.PasteSpecial(Link: true, DisplayAsIcon: true);
        }


    }

   public class Account
    {
        public int ID
        {
            get;
            set;
        }
        public double Balance
        {
            get;
            set;
        }
    }

    
}
