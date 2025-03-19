using System;
using System.IO;
using Microsoft.Reporting.WinForms;

namespace ReportePeriodo
{
    internal class GTHUtils
    {
        public static void SavePDF(ReportViewer viewer, string savePath)
        {
            var deviceInfo = @"<DeviceInfo>
            <EmbedFonts>Arial</EmbedFonts>
            </DeviceInfo>";
            try
            {
                byte[] Bytes = viewer.LocalReport.Render("PDF", deviceInfo);

                using (FileStream stream = new FileStream(savePath, FileMode.Create))
                {
                    stream.Write(Bytes, 0, Bytes.Length);
                }
            }
            catch (LocalProcessingException ex) { Console.WriteLine(ex.Message); }
            catch (ReportViewerException ex) { Console.WriteLine(ex.Message); }
            catch (IOException ex) { Console.WriteLine(ex.Message); }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        public static void DeleteFile(string ruta)
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(ruta);
            try
            {
                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
