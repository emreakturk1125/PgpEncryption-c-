using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataTable = System.Data.DataTable;

namespace pgp
{
    class Program
    {
        static void Main(string[] args)
        {

            string fileLocation = @"C:\Users\emre.akturk\Desktop\bilin\ŞifrelemeKlasör";

            if (!Directory.Exists(fileLocation))
            {
                Directory.CreateDirectory(fileLocation);
            }

            StreamWriter writer;
            FileStream ostrm;
            string fileLocationInput = @"C:\Users\emre.akturk\Desktop\bilin\ŞifrelemeKlasör\inputFile.txt";
            //string fileLocationEncrytedInput = @"C:\Users\emre.akturk\Desktop\bilin\ŞifrelemeKlasör\ReadEncryptedInput.pgp";
            string fileLocationEncrytedInput = @"C:\Users\emre.akturk\Desktop\bilin\ŞifrelemeKlasör\EncryptedInputFile.txt";
            string fileLocationDecryptedOutput = @"C:\Users\emre.akturk\Desktop\bilin\ŞifrelemeKlasör\DecryptedOutputFile.txt";

            if (!File.Exists(fileLocationInput) && !File.Exists(fileLocationEncrytedInput) && !File.Exists(fileLocationDecryptedOutput))
            {
                using (ostrm = File.Create(fileLocationInput)) { }
                using (ostrm = File.Create(fileLocationEncrytedInput)) { }
                using (ostrm = File.Create(fileLocationDecryptedOutput)) { }
            }
            //else
            //{
            //    File.Delete(fileLocationInput);
            //    File.Delete(fileLocationEncrytedInput);
            //    File.Delete(fileLocationDecryptedOutput);
            //}

            TextWriter oldOut = Console.Out;
            try
            {
                ostrm = new FileStream(fileLocationInput, FileMode.OpenOrCreate, FileAccess.Write);
                writer = new StreamWriter(ostrm);
            }
            catch (Exception e)
            {
                Console.WriteLine("Cannot open input.txt for writing");
                Console.WriteLine(e.Message);
                return;
            }


            Console.WriteLine("Şifrelemek istediğiniz metni giriniz : ");

            Console.SetOut(writer);
            string yazı = Console.ReadLine();
            //Console.WriteLine(yazı);
            Console.SetOut(oldOut);
            writer.Close();
            ostrm.Close();
            Console.WriteLine("Done");
            Console.WriteLine("-----------------------------------------------------");
            Console.WriteLine("Şifrelemek istiyormusun ?(E/H)");
            Char cevap = Convert.ToChar(Console.ReadLine());

            if (cevap == 'E' || cevap == 'e')
            {
                PGPEncryptDecrypt.EncryptAndSign(fileLocationInput, fileLocationEncrytedInput, @"C:\Users\emre.akturk\Desktop\kalın\kalin_public_key.asc", @"C:\Users\emre.akturk\Desktop\bilin\bilin_private_key.asc", "bilinyazilim34", true);
                Console.WriteLine("Encryted..");
                string line;
                try
                {
                    StreamReader sr = new StreamReader(fileLocationEncrytedInput);
                    line = sr.ReadLine();

                    while (line != null)
                    {
                        Console.WriteLine(line);
                        line = sr.ReadLine();
                    }

                    sr.Close(); 
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception: " + e.Message);
                }
            }

            Console.WriteLine("-----------------------------------------------------");
            Console.WriteLine("Şifreyi çözmek  istiyormusun ?(E/H)");
            Char cevap2 = Convert.ToChar(Console.ReadLine());

            if (cevap == 'E' || cevap == 'e')
            {
                PGPEncryptDecrypt.Decrypt(fileLocationEncrytedInput, @"C:\Users\emre.akturk\Desktop\kalın\kalin_private_key.asc", "kalinyazilim34", fileLocationDecryptedOutput);
                Console.WriteLine("Decrypted..");

                string line;
                try
                {
                    StreamReader sr = new StreamReader(fileLocationDecryptedOutput);
                    line = sr.ReadLine();

                    while (line != null)
                    {
                        Console.WriteLine(line);
                        line = sr.ReadLine();
                    }

                    sr.Close();
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception: " + e.Message);
                }
            }
            Console.WriteLine("-----------------------------------------------------");
            Console.WriteLine("Verileri  xlx  ve   csv formarlarına aktarmak istiyormusun(e/h)");
            string cvp = Convert.ToString(Console.ReadLine());

            if (cvp == "e")
            {
                string path = @"C:\Users\emre.akturk\Desktop\bilin\ŞifrelemeKlasör\inputFile.txt";
                DataTable table = ReadFile(path);
                Excel_FromDataTable(table);
                Console.WriteLine("Excele aktarıldı");
            }

            Console.ReadKey();

        }

        private static void Excel_FromDataTable(DataTable dt)
        {
            // EXCELE AKTAR - START

            Application excel = new Application();
            Workbook workbook = excel.Application.Workbooks.Add(true);  

            int iCol = 0;
            foreach (DataColumn c in dt.Columns)
            {
                iCol++;
                excel.Cells[1, iCol] = c.ColumnName;
            }
            int iRow = 0;
            foreach (DataRow r in dt.Rows)
            {
                iRow++;

                iCol = 0;
                foreach (DataColumn c in dt.Columns)
                {
                    iCol++;
                    excel.Cells[iRow + 1, iCol] = r[c.ColumnName];
                }
            }

            Worksheet myWorkSheet = (Worksheet)workbook.Worksheets.get_Item(1);
            
            Range range = (Range)myWorkSheet.Application.Rows[1, Type.Missing];
            range.Select();
            range.Delete(XlDirection.xlUp);

            object missing = System.Reflection.Missing.Value;

            workbook.SaveAs(@"C:\Users\emre.akturk\Desktop\bilin\ŞifrelemeKlasör\inputFile(" + Guid.NewGuid() + ").xls",
                XlFileFormat.xlXMLSpreadsheet, missing, missing,
                false, false, XlSaveAsAccessMode.xlNoChange,
                missing, missing, missing, missing, missing);

            excel.Visible = true;
            Worksheet worksheet = (Worksheet)excel.ActiveSheet;
            ((_Worksheet)worksheet).Activate();

            ((_Application)excel).Quit();

            //EXCELE AKTAR - END

            // CSV DOSYASINA AKTAR - START

            DataTable tbl = dt;

            string filePath = @"C:\Users\emre.akturk\Desktop\bilin\ŞifrelemeKlasör\test.csv";
            string delimiter = ",";

            StringBuilder sb = new StringBuilder();
            List<string> CsvRow = new List<string>();


            //foreach (DataColumn c in dt.Columns)
            //{
            //    CsvRow.Add(c.ColumnName);
            //}
            //sb.AppendLine(string.Join(delimiter, CsvRow));

            foreach (DataRow r in dt.Rows)
            {
                CsvRow.Clear();

                foreach (DataColumn c in dt.Columns)
                {
                    CsvRow.Add(r[c].ToString());
                }

                sb.AppendLine(string.Join(delimiter, CsvRow));
            }

            File.WriteAllText(filePath, sb.ToString());



            //CSV DOSYASINA AKTAR - END
        }

        private static DataTable ReadFile(string path)
        {
            System.Data.DataTable table = new System.Data.DataTable("dataFromFile");
            table.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("col1", typeof(string))
            });
            using (StreamReader sr = new StreamReader(path))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {

                    DataRow tempRw = table.NewRow();
                    tempRw["col1"] = line;
                    table.Rows.Add(tempRw);
                }
            }
            return table;
        }

    }

}
