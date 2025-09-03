using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        //#if DEBUG
        //        args = new string[4];
        //        args[0] = @"C:\Users\mazurrad\source\repos\ExcelToCsvConverter\test\plik.xlsx"; // Excel
        //        args[1] = @"C:\Users\mazurrad\source\repos\ExcelToCsvConverter\test\plik.csv";  // CSV
        //        args[2] = @";"; // Separator
        //        args[3] = @"1"; // Indeks skoroszytu
        //#endif

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        try
        {
            Console.WriteLine("Program wystartował");
            // Walidacja argumentów
            if (args.Length < 4)
            {
                Console.WriteLine("Błąd: Nie podano wszystkich wymaganych argumentów: [ExcelPath] [CSVPath] [Separator] [WorksheetIndex]");
                Environment.Exit(1);
            }

            //Console.WriteLine("KROK 1");

            string excelPath = args[0].Trim();
            string csvPath = args[1].Trim();
            string separator = args[2].Trim();


            //Console.WriteLine("KROK 2");

            if (string.IsNullOrWhiteSpace(excelPath) || string.IsNullOrWhiteSpace(csvPath))
            {
                Console.WriteLine("Błąd: Ścieżka do pliku Excel lub CSV jest pusta.");
                Environment.Exit(1);
            }


            // Console.WriteLine("KROK 3");

            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Błąd: Plik Excel nie istnieje: {excelPath}");
                Environment.Exit(1);
            }


            // Console.WriteLine("KROK 4");

            if (!int.TryParse(args[3], out int worksheetIndex))
            {
                Console.WriteLine($"Błąd: Niepoprawny indeks skoroszytu: {args[3]}");
                Environment.Exit(1);
            }


            // Console.WriteLine("KROK 5");

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                if (worksheetIndex < 0 || worksheetIndex >= package.Workbook.Worksheets.Count)
                {
                    Console.WriteLine($"Błąd: Indeks skoroszytu {worksheetIndex} jest poza zakresem. Liczba dostępnych skoroszytów: {package.Workbook.Worksheets.Count}");
                    Environment.Exit(1);
                }



                //  Console.WriteLine("KROK 6");

                var worksheet = package.Workbook.Worksheets[worksheetIndex];

                if (worksheet.Dimension == null)
                {
                    Console.WriteLine($"Błąd: Wybrany skoroszyt ({worksheetIndex}) jest pusty.");
                    Environment.Exit(1);
                }



                //Console.WriteLine("KROK 7");

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                using (var writer = new StreamWriter(csvPath))
                {
                    for (int row = 1; row <= rowCount; row++)
                    {
                        string[] rowValues = new string[colCount];

                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = worksheet.Cells[row, col].Text ?? string.Empty;

                            // Podwajanie cudzysłowów i zamykanie w cudzysłowie, jeśli potrzebne
                            cellValue = cellValue.Replace("\"", "\"\"");

                            if (cellValue.Contains(separator) || cellValue.Contains("\"") || cellValue.Contains("\n") || cellValue.Contains(" "))
                                cellValue = $"\"{cellValue}\"";

                            rowValues[col - 1] = cellValue;
                        }

                        writer.WriteLine(string.Join(separator, rowValues));
                    }
                }
            }

            Console.WriteLine("Plik został poprawnie przekonwertowany do CSV!");
        }
        catch (UnauthorizedAccessException)
        {
            Console.WriteLine("Błąd: Brak uprawnień do odczytu/zapisu pliku.");
            Environment.Exit(1);
        }
        catch (IOException ioEx)
        {
            Console.WriteLine($"Błąd pliku: {ioEx.Message}");
            Environment.Exit(1);
        }
        catch (FormatException)
        {
            Console.WriteLine("Błąd: Nieprawidłowy format liczbowy (indeks skoroszytu).");
            Environment.Exit(1);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Nieoczekiwany błąd: {ex.Message}");
            Environment.Exit(1);
        }
    }
}
