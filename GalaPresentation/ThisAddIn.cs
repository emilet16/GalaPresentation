using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Collections;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;

namespace GalaPresentation
{
    public partial class ThisAddIn
    {
        private PowerPoint.Presentation Pres;

        private void generateFiles(Excel.Workbook workbook, string picturesDir)
        {
            var layouts = Pres.SlideMaster.CustomLayouts;

            for (int i = 1; i <= workbook.Sheets.Count; i++)
            {
                var worksheet = workbook.Sheets[i] as Excel.Worksheet;
                var title = worksheet.Name;
                var usedRange = worksheet.UsedRange;

                List<string[]> winners = extractData(usedRange);

                List<List<string[]>> chunks = sortIntoChunks(winners);

                foreach (var chunk in chunks)
                {
                    var length = chunk.Count;
                    var layout = layouts[length + 1];

                    Pres.Slides.AddSlide(Pres.Slides.Count + 1, layout);
                    var slide = Pres.Slides[Pres.Slides.Count];
                    slide.Shapes.Title.TextFrame.TextRange.Text = title;

                    List<PowerPoint.Shape> nameBoxes = new List<PowerPoint.Shape>();
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                        {
                            nameBoxes.Add(shape);
                        }
                    }

                    int j = 0;
                    foreach (var winner in chunk)
                    {
                        var name = winner[0];
                        var imagePath = picturesDir + "\\" + ((winner[1].Length > 4) ? "P00" : "P000") + winner[1] + ".jpg";
                        slide.Shapes.AddPicture(imagePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
                        nameBoxes[j].TextFrame.TextRange.Text = name;
                        j++;
                    }
                }
            }
        }

        public void run()
        {
            MessageBox.Show("Veuillez choisir le fichier Excel contenant les gagnants.");

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(openFileDialog.FileName))
                {
                    string filePath = openFileDialog.FileName;

                    MessageBox.Show("Veuillez choisir le dossier contenant les photos des élèves.");

                    using (var fbd = new FolderBrowserDialog())
                    {
                        if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                        {
                            string picturesDir = fbd.SelectedPath;
                            var ExcelApp = new Excel.Application();
                            var workbook = ExcelApp.Workbooks.Open(filePath);

                            generateFiles(workbook, picturesDir);

                            workbook.Close();
                            ExcelApp.Quit();
                        }
                    }
                }
            }
        }

        private List<string[]> extractData(Excel.Range range)
        {
            List<string[]> output = new List<string[]>();
            var size = range.Rows.Count;

            for (int j = 1; j <= size; j++)
            {
                var name = range.Cells[j, 1].Text as string;
                var id = range.Cells[j, 2].Text as string;
                if (name == "" || id == "") continue;
                string[] data = { name, id };
                output.Add(data);
            }

            return output;
        }

        private List<List<string[]>> sortIntoChunks(List<string[]> array)
        {
            var size = array.Count;
            var numChunks = Decimal.ToInt32(Decimal.Truncate((size - 1) / 20) + 1);
            var minChunkSize = Decimal.ToInt32(Decimal.Truncate(size / numChunks));
            var numSmallChunks = numChunks * (minChunkSize + 1) - size;

            List<List<string[]>> chunks = new List<List<string[]>>();

            for (int j = 0; j < numChunks; j++)
            {
                if (j < numSmallChunks)
                {
                    chunks.Add(array.GetRange(0, minChunkSize));
                    array.RemoveRange(0, minChunkSize);
                }
                else
                {
                    chunks.Add(array.GetRange(0, minChunkSize + 1));
                    array.RemoveRange(0, minChunkSize + 1);
                }
            }

            return chunks;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.PresentationOpen += Application_PresentationOpen;
        }

        private void Application_PresentationOpen(Presentation pres)
        {
            Pres = pres;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
