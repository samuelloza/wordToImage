
/**
* @author Samuel Loza Ramirez
*
* @create date - 2020-16-12 
* @desc Plugin para Word, guarta todas las hojas en formato de imágenes
*/

using Microsoft.Office.Tools.Ribbon;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordToImage
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            object Miss = System.Reflection.Missing.Value;
            object What = Microsoft.Office.Interop.Word.WdGoToItem.wdGoToPage;
            object Which = Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute;
            object Start;
            object End;
            object CurrentPageNumber;
            object NextPageNumber;
            String fileNameDirectory = "C:\\excel\\";

            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Seleccione el directorio donde se guardaran las imágenes";

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                fileNameDirectory = fbd.SelectedPath.Replace(@"\", @"\\") + @"\\";
            }
            else
            {
                MessageBox.Show("Seleccione el directorio donde se guardaran las imágenes ",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);

                return;
            }

            Debug.WriteLine(fileNameDirectory + " ............ ");

            //Obtain Current Document
            Word.Document Doc = Globals.ThisAddIn.Application.ActiveDocument;

            // Get pages count
            Microsoft.Office.Interop.Word.WdStatistic PagesCountStat = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages;
            int PagesCount = Doc.ComputeStatistics(PagesCountStat, ref Miss);


            for (int Index = 1; Index < PagesCount + 1; Index++)
            {
                CurrentPageNumber = (Convert.ToInt32(Index.ToString()));
                NextPageNumber = (Convert.ToInt32((Index + 1).ToString()));

                // Get start position of current page
                Start = Globals.ThisAddIn.Application.Selection.GoTo(ref What, ref Which, ref CurrentPageNumber, ref Miss).Start;

                // Get end position of current page                                
                End = Globals.ThisAddIn.Application.Selection.GoTo(ref What, ref Which, ref NextPageNumber, ref Miss).End;

                saveImage(Doc, Start, End, fileNameDirectory, Index);

            }
            MessageBox.Show("Se creo correctamente las imágenes", "Super");
            Process.Start(fileNameDirectory);

            //  Debug.WriteLine(PagesCount);
        }

        private static void saveImage(Word.Document Doc, Object Start, Object End, String fileNameDirectory, int Index)
        {
            Object bits;

            if (Start.ToString().Equals(End.ToString()))
            {
                //Debug.WriteLine(Doc.Range(ref Start).Text);
                bits = Doc.Range(ref Start).EnhMetaFileBits;
            }
            else
            {
                // Debug.WriteLine(Doc.Range(ref Start, ref End).Text);
                bits = Doc.Range(ref Start, ref End).EnhMetaFileBits;
            }

            try
            {
                //https://stackoverflow.com/questions/20326478/convert-word-file-pages-to-jpg-images-using-c-sharp
                using (var ms = new MemoryStream((byte[])(bits)))
                {
                    var image = System.Drawing.Image.FromStream(ms);
                    var pngTarget = Path.ChangeExtension(fileNameDirectory + String.Format("img_{0}.png", Index), "png");
                    image.Save(pngTarget, System.Drawing.Imaging.ImageFormat.Png);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Algo paso al guardar la imagen",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
               );
            }
        }
    }
}
