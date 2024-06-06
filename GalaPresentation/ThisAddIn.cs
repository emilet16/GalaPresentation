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
using System.Text.RegularExpressions;

namespace GalaPresentation
{
    public partial class ThisAddIn
    {
        private PowerPoint.Presentation Pres;
        int[] slidesLayouts = { 1, 2, 3, 4, 5, 6, 8, 10, 12, 15, 18, 21, 24, 27, 30 };

        private void generateFiles(Excel.Workbook workbook, string picturesDir)
        {
            var layouts = Pres.SlideMaster.CustomLayouts;

            var worksheet = workbook.Sheets[1] as Excel.Worksheet;
            var usedRange = worksheet.UsedRange;

            List<string[]> winners = extractData(usedRange);
            List<List<string[]>> sortedWinners = new List<List<string[]>>();

            foreach(var winner in winners)
            {
                var category = winner[3];
                var categoryExists = false;
                foreach(var winnerCat in sortedWinners)
                {
                    if (category.Equals(winnerCat[0][3]))
                    {
                        winnerCat.Add(winner);
                        categoryExists = true;
                        break;
                    }
                }
                if (!categoryExists)
                {
                    var newCat = new List<string[]> { winner };
                    sortedWinners.Add(newCat);
                }
            }

            foreach(var winnerCategory in sortedWinners)
            {
                var title = winnerCategory[0][2];

                var sortOutput = sortIntoChunks(winnerCategory);
                List<List<string[]>> chunks = sortOutput.Item1;
                int largestSize = sortOutput.Item2;
                int closestBiggerLayout = slidesLayouts.Aggregate((x,y)=>(Math.Abs(x-largestSize) < Math.Abs(y-largestSize))&&(x>=largestSize) ? x : y);
                int index = Array.IndexOf(slidesLayouts, closestBiggerLayout) + 1;

                foreach (var chunk in chunks)
                {
                    var layout = layouts[index];

                    var slide = Pres.Slides.AddSlide(Pres.Slides.Count + 1, layout);

                    List<PowerPoint.Shape> nameBoxes = new List<PowerPoint.Shape>();

                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                        {
                            nameBoxes.Add(shape);
                        }
                    }

                    nameBoxes[0].TextFrame.TextRange.Text = title;

                    int j = 1;
                    foreach (var winner in chunk)
                    {
                        var name = winner[0];
                        try
                        {
                            var imagePath = picturesDir + "\\" + winner[1] + ".jpg";
                            slide.Shapes.AddPicture(imagePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
                        }
                        catch
                        {
                            try
                            {
                                var imagePath = picturesDir + "\\" + ((winner[1].Length > 4) ? "P00" : "P000") + winner[1] + ".jpg";
                                slide.Shapes.AddPicture(imagePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 0, 0);
                            }
                            catch
                            {
                                MessageBox.Show("Il n'y a pas pas de photos disponible pour l'élève " + name + "(" + winner[1] + ")" + " pour " + title);
                                continue;
                            }
                        }
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
                var id = range.Cells[j, 1].Text as string;
                id = Regex.Replace(id, @"\s+", "");
                int res;
                if (!Int32.TryParse(id, out res)) continue;
                var firstName = range.Cells[j, 2].Text as string;
                var lastName = range.Cells[j, 3].Text as string;
                var name = firstName + " " + lastName;
                var title = range.Cells[j, 5].Text as string;
                var category = range.Cells[j, 6].Text as string;
                category = Regex.Replace(title+category, @"\s+", "");
                if (name == "" || id == "" || title == "" || category == "") continue;
                string[] data = { name, id, title, category };
                output.Add(data);
            }

            return output;
        }

        private (List<List<string[]>>, int largestSize) sortIntoChunks(List<string[]> array)
        {
            var size = array.Count;
            var numChunks = Decimal.ToInt32(Decimal.Truncate((size - 1) / 30) + 1);
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
            int largestSize = numChunks == numSmallChunks ? minChunkSize : minChunkSize + 1;

            return (chunks, largestSize);
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
