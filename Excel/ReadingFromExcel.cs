using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    internal class ReadingFromExcel : IReading
    {
        private List<Vagt> Vagter;

        public ReadingFromExcel()
        {
            Vagter = new List<Vagt>();
        }


        public List<Vagt> ReadData(string target, string FilePath)
        {
            using (var workbook = new XLWorkbook(FilePath))
            {
                // Specify the worksheet where you want to search
                IXLWorksheet worksheet = workbook.Worksheet(1); // Change 1 to the appropriate worksheet index or name

                // Define the name you want to search for
            
                List<string> cellAddress = new List<string>();
                List<string> weekDays = new List<string>()
            {
                "Mandag",
                "Tirsdag",
                "Onsdag",
                "Torsdag",
                "Fredag",
                "Lørdag",
                "Søndag"
            };

                string Dag = "";
                string dateString = "";
                string timeString = "";


                // Loop through all the cells in the worksheet
                foreach (IXLCell cell in worksheet.CellsUsed())
                {
                    // Check if the cell contains the target name
                    if (cell.Value.ToString().Equals(target, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Found '{target}' at cell {cell.Address}");
                        cellAddress.Add(cell.Address.ToString());
                        // If you want to stop searching after the first occurrence, you can break here.
                    }
                }

                //todo: find if there is a time on the top of the name
                string adressabove;
                foreach (string address in cellAddress)
                {

                    IXLCell cellAbove = worksheet.Cell(address).CellAbove();
                    if (cellAbove != null)
                    {
                        Console.WriteLine($"Data above '{target}' at cell {address}: {cellAbove.Value}");
                        timeString = cellAbove.Value.ToString();
                    }
                    while (!weekDays.Contains(cellAbove.Value.ToString()))
                    {
                        adressabove = cellAbove.Address.ToString();
                        cellAbove = worksheet.Cell(adressabove).CellAbove();
                    }
                    if (weekDays.Contains(cellAbove.Value.ToString()))
                    {
                        Console.WriteLine(cellAbove.Value.ToString());
                        Dag = cellAbove.Value.ToString();
                        adressabove = cellAbove.Address.ToString();

                        cellAbove = worksheet.Cell(adressabove).CellAbove();
                        Console.WriteLine(cellAbove.Value.ToString());
                        dateString = cellAbove.Value.ToString();


                    }

           
                    Vagt vagt = new Vagt()
                    {
                        dato = dateString,
                        hours = timeString,
                        dag = Dag

                    };
                    Vagter.Add(vagt);

                }

            }
            return Vagter;

        }
    }
}
