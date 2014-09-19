using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Drawing.Imaging;

namespace tiffTopdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            loadLstFiles();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string pdfFile = null;
            string fileName = txtFolder.Text.Trim();
            if (!fileName.EndsWith("\\"))
                fileName += "\\";

            fileName += lstFiles.SelectedItem.ToString();
            if (File.Exists(fileName))
            {
                pdfFile = ConvertTiffToPdf(fileName);
                if (pdfFile != null)
                {
                    try
                    {
                        string destFile = Environment.CurrentDirectory + "\\data\\";
                        destFile = destFile.Substring(0, destFile.LastIndexOf('\\') + 1);
                        if (Directory.Exists(destFile))
                        {
                            destFile += Path.GetFileName(pdfFile);
                            File.Copy(pdfFile, destFile);
                            File.Delete(pdfFile);
                        }
                    }
                    catch { }
                }
            }

            loadLstFiles();
        }

        private void loadLstFiles()
        {
            lstFiles.Items.Clear();
            
            string currentDir = Environment.CurrentDirectory + "\\data";

            txtFolder.Text = currentDir;

            string[] files = Directory.GetFiles(currentDir);

            foreach (string file in files)
            {
                lstFiles.Items.Add(Path.GetFileName(file));

            }
        }

        private string ConvertTiffToPdf(string tiffFileName)
        {
            //iTextSharp.text.Rectangle pgSize = PageSize.LETTER;
            iTextSharp.text.Rectangle pgSize = PageSize.LEGAL;
            string pdfFile;

            do
            {
                pdfFile = Path.GetDirectoryName(tiffFileName) + Path.DirectorySeparatorChar +
                    "Signed_Investor_AFE_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".pdf";
            } while (File.Exists(pdfFile));

            if (File.Exists(pdfFile))
                File.Delete(pdfFile);

            if (!File.Exists(tiffFileName))
                return null;

            Bitmap bmp = new Bitmap(tiffFileName);

            int totalPages = bmp.GetFrameCount(FrameDimension.Page);

            // creation of the document with a certain size and certain margins
            float margin = 0.0f;
            //float margin = 50.0f;
            Document document = new Document(pgSize, margin, margin, margin, margin);

            try
            {

                // creation of the different writers
                PdfWriter writer = PdfWriter.GetInstance(document,
                    new FileStream(pdfFile, FileMode.Create));

                // Which of the multiple images in the TIFF file do we want to load
                // 0 refers to the first, 1 to the second and so on.
                document.Open();
                PdfContentByte cb = writer.DirectContent;
                iTextSharp.text.Image img;

                for (int ii = 0; ii < totalPages; ii++)
                {
                    bmp.SelectActiveFrame(FrameDimension.Page, ii);
                    img = iTextSharp.text.Image.GetInstance(BitmapToBytes(bmp));
                    img.ScalePercent(72f / bmp.HorizontalResolution * 100);
                    img.SetAbsolutePosition(0, 0);
                    cb.AddImage(img);
                    document.NewPage();
                }

                document.Close();
                document = null;
                bmp.Dispose();
                bmp = null;

            }
            catch (Exception ex)
            {
                this.Cursor = Cursors.Default;
                MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            return pdfFile;
        }

        private byte[] BitmapToBytes(Bitmap bmp)
        {
            byte[] data = new byte[0];
            using (MemoryStream ms = new MemoryStream())
            {
                bmp.Save(ms, ImageFormat.Png);
                ms.Seek(0, 0);
                data = ms.ToArray();
            }
            return data;
        }

        private void lstFiles_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            string path = txtFolder.Text.Trim();
            while (path.EndsWith("\\"))
                path = path.Substring(0, path.Length - 1).Trim();

            string fname = path + "\\" + lstFiles.SelectedItem.ToString();

            try
            {
                System.Diagnostics.Process.Start(fname);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

    }
}
