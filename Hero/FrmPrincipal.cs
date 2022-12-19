using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Web;
using System.IO;
using iTextSharp.text.pdf;
using System.Linq;
using LovePdf.Core;
using LovePdf.Model.Task;
using LovePdf.Model.TaskParams;
using LovePdf.Model.Enums;
using System.Threading;


namespace Baymax
{
    public partial class Baymax : Form
    {
        string dir;
        string destino;
        string salidaDocumentos;
        string tempPath = System.IO.Path.GetTempPath();

        public Baymax()
        {
            InitializeComponent();
            chLaMejor.Checked = true;
          
            ch0.Checked = true;
        }


    //.................................Comprimir Archivo PDF ...................................//
        private void btnSeleccionar_Click(object sender, EventArgs e)
        {
            //String file_name = string.Empty;

            //this.openFileDialog1.Filter ="Images (*.PDF)|*.PDF|" +"All files (*.*)|*.*";
            //this.openFileDialog1.Title = "Seleccionar Archivo Pdf";

            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
            //{
            //    dir = openFileDialog1.FileName;
            //    destino = Path.GetFileName(dir);
            //    lsDocComprimir.Text = destino;

            //    FileInfo info = new FileInfo(dir);
            //    long value = info.Length;
            //    if (value / 1024 <= 1024) //kb
            //    {
            //        lbTamanioA.Text = string.Format("{0:N0}", (value / 1024f)) + " Kb";
            //    }
            //    if (value / 1024 > 1024) //Mb
            //    {
            //        lbTamanioA.Text = string.Format("{0:N0}", (value / 1024f)) + " Mb";
            //    }
            //}

            String file_name = string.Empty;

            this.openFileDialog3.Filter =
           "Images (*.PDF)|*.PDF|" +
           "All files (*.*)|*.*";
            this.openFileDialog5.Title = "Seleccionar Archivo Pdf";

            if (openFileDialog5.ShowDialog() == DialogResult.OK)
                lsDocComprimir.Items.AddRange(openFileDialog5.FileNames);
        }

        private void btnComprimir_Click(object sender, EventArgs e)
        {
            try
            {

                if (chBuena.Checked)
                {
                   
                    MessageBoxTemporal.Show("Espere mientras se Comprime el PDF.", " Baymax v1.5", 3, true);
                    foreach (string arquivoOrigem1 in lsDocComprimir.Items.Cast<string>().ToArray())
                    {
                        //string pdfFile = dir;
                        //PdfReader reader = new PdfReader(pdfFile);
                        //PdfStamper stamper = new PdfStamper(reader, new FileStream(@"D:\" + destino, FileMode.Create), PdfWriter.VERSION_1_5);
                        //stamper.FormFlattening = true;
                        //stamper.Close();
                        //txtDocComprimir.Clear();
                        //MessageBoxTemporal.Show("Espere mientras el archivo se reduce de tamaño", " Baymax v1.4", 5, false);
                        //FileInfo info = new FileInfo(@"D:\" + destino);
                        //long value = info.Length;
                        //if (value / 1024 <= 1024) //kb
                        //{
                        //    lbTamnioR.Text = string.Format("{0:N0}", (value / 1024f)) + " Kb";
                        //}
                        //if (value / 1024 > 1024) //Mb
                        //{
                        //    lbTamnioR.Text = string.Format("{0:N0}", (value / 1024f)) + " Mb";
                        //}

                        var api = new LovePdfApi("project_public_681de94b8592545e00e8ea3aa6b85cef_zjS6D000863c42d932d8966da0023c96b8731", "secret_key_f9b5c586555f47c6b0a1c542db712ba9_8qTgT48d1dd3f7a50c806bcd52fe11aa51328");

                        //create compress task
                        var task = api.CreateTask<CompressTask>();

                        //file variable contains server file name


                        if (ch0.Checked)
                        {
                            var file = task.AddFile(arquivoOrigem1, task.TaskId, Rotate.Degrees0);
                        }
                        if (ch90.Checked)
                        {
                            var file = task.AddFile(arquivoOrigem1, task.TaskId, Rotate.Degrees90);
                        }
                        if (ch180.Checked)
                        {
                            var file = task.AddFile(arquivoOrigem1, task.TaskId, Rotate.Degrees180);
                        }
                        if (ch270.Checked)
                        {
                            var file = task.AddFile(arquivoOrigem1, task.TaskId, Rotate.Degrees270);
                        }

                        //proces added files
                        //time var will contains information about time spent in process
                        var time = task.Process();

                        //download files to specific folder
                        task.DownloadFile(salidaDocumentos+"\\");

                        //FileInfo info = new FileInfo(@"D:\" + Path.GetFileName(arquivoOrigem1));

                        //long value = info.Length;
                        //if (value / 1024 <= 1024) //kb
                        //{
                        //    lbTamnioR.Text = string.Format("{0:N0}", (value / 1024f)) + " Kb";
                        //}
                        //if (value / 1024 > 1024) //Mb
                        //{
                        //    lbTamnioR.Text = string.Format("{0:N0}", (value / 1024f)) + " Mb";
                        //}
                    }
                    
                    MessageBox.Show("Se ha Reducido el Documento/s\nBusque su Archivo", " Baymax v1.5", MessageBoxButtons.OK, MessageBoxIcon.Asterisk); 
                    lsDocComprimir.Items.Clear();
                }

                else
                {
                    foreach (string arquivoOrigem1 in lsDocComprimir.Items.Cast<string>().ToArray())
                    {
                        //Aspose.Pdf.Document document = new Document(dir);
                        //// Optimzie the file size by removing unused objects
                        //document.OptimizeResources(new Document.OptimizationOptions()
                        //{
                        //    LinkDuplcateStreams = true,
                        //    RemoveUnusedObjects = true,
                        //    RemoveUnusedStreams = true,
                        //    CompressImages = true,
                        //    ImageQuality = 10
                        //});
                        //// Save the updated file
                        //document.Save(@"D:\" + destino);
                        //txtDocComprimir.Clear();
                        //MessageBoxTemporal.Show("Espere mientras el archivo se reduce de tamaño", " Baymax v1.4", 5, false);


                        var api = new LovePdfApi("project_public_681de94b8592545e00e8ea3aa6b85cef_zjS6D000863c42d932d8966da0023c96b8731", "secret_key_f9b5c586555f47c6b0a1c542db712ba9_8qTgT48d1dd3f7a50c806bcd52fe11aa51328");

                        var task = api.CreateTask<CompressTask>();


                        //add file, and specify rotation
                        MessageBoxTemporal.Show("Espere mientras se Comprime el PDF", " Baymax v1.5", 3, false);

                        if (ch0.Checked)
                        {
                            var file = task.AddFile(arquivoOrigem1, task.TaskId, Rotate.Degrees0);
                        }
                        if (ch90.Checked)
                        {
                            var file = task.AddFile(arquivoOrigem1, task.TaskId, Rotate.Degrees90);
                        }
                        if (ch180.Checked)
                        {
                            var file = task.AddFile(arquivoOrigem1, task.TaskId, Rotate.Degrees180);
                        }
                        if (ch270.Checked)
                        {
                            var file = task.AddFile(arquivoOrigem1, task.TaskId, Rotate.Degrees270);
                        }

                        //set compress parameters and process files
                        var time = task.Process(new CompressParams
                        {
                            CompressionLevel = CompressionLevels.Extreme,
                            OutputFileName = "extreme_compression"
                        });

                        //download output file(s) to specific directory
                        task.DownloadFile(@salidaDocumentos+"\\");

                        //FileInfo info = new FileInfo(@"D:\" + Path.GetFileName(arquivoOrigem1));

                        //long value = info.Length;
                        //if (value / 1024 <= 1024) //kb
                        //{
                        //    lbTamnioR.Text = string.Format("{0:N0}", (value / 1024f)) + " Kb";
                        //}
                        //if (value / 1024 > 1024) //Mb
                        //{
                        //    lbTamnioR.Text = string.Format("{0:N0}", (value / 1024f)) + " Mb";
                        //}
                    }
                   
                    MessageBox.Show("Se ha Reducido el Documento/s\nBusque su Archivo", " Baymax v1.5", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    lsDocComprimir.Items.Clear();
                }
            }  
            catch
            {
                MessageBox.Show(e.ToString());
                MessageBox.Show("Ha Sucedido un Error Intente de Nuevo"); }
        }

        private void btnSeleccionar_MouseHover(object sender, EventArgs e)
        {
            btnSeleccionar.BackColor = System.Drawing.Color.White;
            btnSeleccionar.ForeColor = System.Drawing.Color.Black;
        }

        private void btnSeleccionar_MouseLeave(object sender, EventArgs e)
        {
            btnSeleccionar.BackColor = System.Drawing.Color.Green;
            btnSeleccionar.ForeColor = System.Drawing.Color.White;
        }

        private void btnSelecDesproteger_MouseHover(object sender, EventArgs e)
        {
            btnSelecDesproteger.BackColor = System.Drawing.Color.White;
            btnSelecDesproteger.ForeColor = System.Drawing.Color.Black;
        }

        private void btnSelecDesproteger_MouseLeave(object sender, EventArgs e)
        {
            btnSelecDesproteger.BackColor = System.Drawing.Color.Green;
            btnSelecDesproteger.ForeColor = System.Drawing.Color.White;
        }

        private void btnDesproteger_MouseHover(object sender, EventArgs e)
        {
            btnDesproteger.BackColor = System.Drawing.Color.Red;
            btnDesproteger.ForeColor = System.Drawing.Color.White;
        }

        private void btnDesproteger_MouseLeave(object sender, EventArgs e)
        {
            btnDesproteger.BackColor = System.Drawing.Color.Orange;
            btnDesproteger.ForeColor = System.Drawing.Color.Black;
        }

        private void btnComprimir_MouseHover(object sender, EventArgs e)
        {
            btnComprimir.BackColor = System.Drawing.Color.Red;
            btnComprimir.ForeColor = System.Drawing.Color.White;
        }

        private void btnComprimir_MouseLeave(object sender, EventArgs e)
        {
            btnComprimir.BackColor = System.Drawing.Color.Orange;
            btnComprimir.ForeColor = System.Drawing.Color.Black;
        }
        /*****************************************Comprimir Archivo PDF************************************************/

        //.........................................Convertir Imagen a PDF......................................//
        private void btnImage_Click(object sender, EventArgs e)
        {

            String file_name = string.Empty;

            this.openFileDialog2.Filter =
           "Images (*.BMP;*.JPG;*.GIF)|*.BMP;*.JPG;*.GIF|" +
            "Documentos(*.doc; *.docx;)| *.doc; *.docx; | " +
            "All files (*.*)|*.*";
            this.openFileDialog2.Title = "Seleccionar Archivo Pdf";

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
                lsImagenes.Items.AddRange(openFileDialog2.FileNames);
        }

        private void btnPdf_Click(object sender, EventArgs e)
        {
            MessageBoxTemporal.Show("Espere mientras se Convierte la Imagen a PDF", " Baymax v1.5", 3, false);
            if(!chSeparados.Checked)
            {
                foreach (string arquivoOrigem1 in lsImagenes.Items.Cast<string>().ToArray())
                {
                    try
                    {
                        using (var document = new iTextSharp.text.pdf.PdfDocument())
                        {
                            iTextSharp.text.Rectangle pageSize = null;

                            using (var srcImage = new Bitmap(arquivoOrigem1))
                            {
                                pageSize = new iTextSharp.text.Rectangle(0, 0, srcImage.Width, srcImage.Height);
                            }
                            using (var ms = new MemoryStream())
                            {
                                var Document = new iTextSharp.text.Document(pageSize, 0, 0, 0, 0);
                                iTextSharp.text.pdf.PdfWriter.GetInstance(Document, ms).SetFullCompression();
                                Document.Open();
                                var image = iTextSharp.text.Image.GetInstance(arquivoOrigem1);
                                Document.Add(image);
                                Document.Close();

                                //File.WriteAllBytes(@"D:\" + Path.GetFileName(arquivoOrigem1) + ".pdf", ms.ToArray());

                                File.WriteAllBytes(@tempPath + "\\" + Path.GetFileName(arquivoOrigem1) + ".pdf", ms.ToArray());
                                lsImagenes.Items.Clear();
                            }
                        }
                        //var api = new LovePdfApi("project_public_681de94b8592545e00e8ea3aa6b85cef_zjS6D000863c42d932d8966da0023c96b8731", "secret_key_f9b5c586555f47c6b0a1c542db712ba9_8qTgT48d1dd3f7a50c806bcd52fe11aa51328");

                        ////create unlock task
                        //var task = api.CreateTask<ImageToPdfTask>();

                        ////file variable contains server file name
                        //// set the password witch the document is locked
                        //var file = task.AddFile(arquivoOrigem1, task.TaskId, "test");

                        ////proces added files
                        ////time var will contains information about time spent in process
                        //var time = task.Process();
                        //task.DownloadFile(@"D:\");
                        tempPath = System.IO.Path.GetTempPath();
                        lstArquivosOrigem.Items.Add(@tempPath + "\\" + Path.GetFileName(arquivoOrigem1) + ".pdf");
                    }

                    catch
                    {
                        MessageBox.Show("Ha Sucedido un Error Intente de Nuevo");
                    }
                }


                try
                {

                    string arquivoPDFDestino = salidaDocumentos + "//DocumentoUnido.pdf";//Le Doy la Direccion Fija D:
                    Merge(
                            lstArquivosOrigem.Items.Cast<string>().ToArray(),
                            arquivoPDFDestino);
                    MessageBox.Show(String.Format(
                        "Arquivo {0} Se Han Unido los Pdfs Correctamente", arquivoPDFDestino));
                    

                    lstArquivosOrigem.Items.Clear();
                }
                catch
                {
                    MessageBox.Show("Ha Sucedido un Error Intente de Nuevo");
                }


                MessageBox.Show("Se ha Convertido la Imagen  a Formato PDF Correctamente", " Baymax v1.5", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            }
            else
            {
                foreach (string arquivoOrigem1 in lsImagenes.Items.Cast<string>().ToArray())
                {
                    try
                    {
                        using (var document = new iTextSharp.text.pdf.PdfDocument())
                        {
                            iTextSharp.text.Rectangle pageSize = null;

                            using (var srcImage = new Bitmap(arquivoOrigem1))
                            {
                                pageSize = new iTextSharp.text.Rectangle(0, 0, srcImage.Width, srcImage.Height);
                            }
                            using (var ms = new MemoryStream())
                            {
                                var Document = new iTextSharp.text.Document(pageSize, 0, 0, 0, 0);
                                iTextSharp.text.pdf.PdfWriter.GetInstance(Document, ms).SetFullCompression();
                                Document.Open();
                                var image = iTextSharp.text.Image.GetInstance(arquivoOrigem1);
                                Document.Add(image);
                                Document.Close();

                                //File.WriteAllBytes(@"D:\" + Path.GetFileName(arquivoOrigem1) + ".pdf", ms.ToArray());

                                File.WriteAllBytes(@salidaDocumentos + "\\" + Path.GetFileName(arquivoOrigem1) + ".pdf", ms.ToArray());
                                lsImagenes.Items.Clear();
                            }
                        }
                        //var api = new LovePdfApi("project_public_681de94b8592545e00e8ea3aa6b85cef_zjS6D000863c42d932d8966da0023c96b8731", "secret_key_f9b5c586555f47c6b0a1c542db712ba9_8qTgT48d1dd3f7a50c806bcd52fe11aa51328");

                        ////create unlock task
                        //var task = api.CreateTask<ImageToPdfTask>();

                        ////file variable contains server file name
                        //// set the password witch the document is locked
                        //var file = task.AddFile(arquivoOrigem1, task.TaskId, "test");

                        ////proces added files
                        ////time var will contains information about time spent in process
                        //var time = task.Process();
                        //task.DownloadFile(@"D:\");
                        tempPath = System.IO.Path.GetTempPath();
                        lstArquivosOrigem.Items.Add(salidaDocumentos + "\\" + Path.GetFileName(arquivoOrigem1) + ".pdf");
                    }

                    catch
                    {
                        MessageBox.Show("Ha Sucedido un Error Intente de Nuevo");
                    }
                }

                MessageBox.Show("Se ha Convertido las Imagen  a Formato PDF Correctamente", " Baymax v1.5", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            }
        }

        private void btnImage_MouseHover(object sender, EventArgs e)
        {
            btnImage.BackColor = System.Drawing.Color.White;
            btnImage.ForeColor = System.Drawing.Color.Black;
        }

        private void btnImage_MouseLeave(object sender, EventArgs e)
        {
            btnImage.BackColor = System.Drawing.Color.Green;
            btnImage.ForeColor = System.Drawing.Color.White;
        }

        private void chLaMejor_Click(object sender, EventArgs e)
        {
            chLaMejor.Checked = true;
            chBuena.Checked = false;

            ch0.Checked = true;
            ch90.Checked = false;
            ch180.Checked = false;
            ch270.Checked = false;
        }

        private void chBuena_Click(object sender, EventArgs e)
        {
            chLaMejor.Checked = false;
            chBuena.Checked = true;

            ch0.Checked = true;
            ch90.Checked = false;
            ch180.Checked = false;
            ch270.Checked = false;
        }

        private void btnPdf_MouseHover(object sender, EventArgs e)
        {
            btnPdf.BackColor = System.Drawing.Color.Red;
            btnPdf.ForeColor = System.Drawing.Color.White;
        }

        private void btnPdf_MouseLeave(object sender, EventArgs e)
        {
            btnPdf.BackColor = System.Drawing.Color.Orange;
            btnPdf.ForeColor = System.Drawing.Color.Black;
        }

        /***************************************************Convertir Imagen a PDF *****************************/



        private void btnAcerca_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Software de Ayuda para Realizar Trabajos con Archivos PDF\n1.- Reduce el Tamaño de tus Archivos Pdf(Compresion: Buena o La Mejor).\n2.- Convierte tus Imágenes a Formato PDF.\n3.- Desprotege los Archivos PDF que no puedes imprimir o copiar.\n4.- Puedes unir los Archivos PDF que quieras.\n5.- Convierte tus Archivos PDF a Word(Soporta Version Demo hasta 22 Hojas).\n6.- Puedes dividir tus Archivos PDF que quieras editar.\nNOTA: Todas las conversiones se las guarda en el disco local D: ", " Baymax v1.5", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }


        //............................................PDF a Word.......................................//
        private void btnPdfW_Click(object sender, EventArgs e)
        {
            string pathToPdf = dir;
            MessageBoxTemporal.Show("Espere mientras se Convierte el Pdf a Word", " Baymax v1.5", 3, false);
            string pathToWord = salidaDocumentos+"\\" + destino + ".doc";

            //Convert PDF file to Word file 
            SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();

            f.OpenPdf(pathToPdf);

            if (f.PageCount > 0)
            {
                int result = f.ToWord(pathToWord);

                //Show Word document 
                if (result == 0)
                {
                    //System.Diagnostics.Process.Start(pathToWord);
                    MessageBox.Show("Se ha Convertido el Documento PDF a Formato Word Correctamente", " Baymax v1.5", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    txtWord.Clear();
                }
            }

            //Document pdfDocument = new Document(dir);
            //MessageBoxTemporal.Show("Espere mientras se Convierte el Pdf a Word", " Baymax v1.1", 3, false);
            //// Save the file into a Microsoft document format
            //pdfDocument.Save("D:/" + destino + ".doc", SaveFormat.Doc); 
            //MessageBox.Show("Se ha Convertido el Documento PDF a Formato Word Correctamente", " Baymax v1.1", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            //Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            //wordDocument = appWord.Documents.Open(@"D:\desktop\xxxxxx.docx");
            //wordDocument.ExportAsFixedFormat(@"D:\desktop\DocTo.pdf", WdExportFormat.wdExportFormatPDF); PDF A WORD
        }

        private void btnPdfW_MouseHover(object sender, EventArgs e)
        {
            btnPdfW.BackColor = System.Drawing.Color.Red;
            btnPdfW.ForeColor = System.Drawing.Color.White;
        }

        private void btnPdfW_MouseLeave(object sender, EventArgs e)
        {
            btnPdfW.BackColor = System.Drawing.Color.Orange;
            btnPdfW.ForeColor = System.Drawing.Color.Black;
        }


        private void btnWord_MouseHover(object sender, EventArgs e)
        {
            btnWord.BackColor = System.Drawing.Color.White;
            btnWord.ForeColor = System.Drawing.Color.Black;
        }

        private void btnWord_MouseLeave(object sender, EventArgs e)
        {
            btnWord.BackColor = System.Drawing.Color.Green;
            btnWord.ForeColor = System.Drawing.Color.White;
        }


        private void btnWord_Click(object sender, EventArgs e)
        {
            String file_name = string.Empty;
            this.openFileDialog4.Filter =
           "Documentos(*.PDF;)| *.PDF; | " +
           "All files (*.*)|*.*";
            this.openFileDialog4.Title = "Seleccionar Documento";

            if (openFileDialog4.ShowDialog() == DialogResult.OK)
            {
                dir = openFileDialog4.FileName;
                destino = Path.GetFileName(dir);
                txtWord.Text = destino;
            }
        }

        /********************************************PDF a Word***********************/



        //..........................................Unir PDF.............................//
        private void btnUnirPdf_Click(object sender, EventArgs e)
        {

            try
            {
                //  this.saveFileDialog1.Filter =
                //"Images (*.PDF)|*.PDF|" +
                // "All files (*.*)|*.*";
                //  this.saveFileDialog1.Title = "Guardar Archivo Pdf";

                //  if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                //  {
                //string arquivoPDFDestino = saveFileDialog1.FileName;//Esto Sirve para Seleccionar donde se Guardara el Documento

                MessageBoxTemporal.Show("Espere mientras se Unen los archivos PDFs", " Baymax v1.5", 3, false);
                string arquivoPDFDestino = salidaDocumentos+"//DocumentoUnido.pdf";//Le Doy la Direccion Fija D:
                Merge(
                        lstArquivosOrigem.Items.Cast<string>().ToArray(),
                        arquivoPDFDestino);
                MessageBox.Show(String.Format(
                    "Arquivo {0} Se Han Unido los Pdfs Correctamente", arquivoPDFDestino));
                //}

                lstArquivosOrigem.Items.Clear();
            }
            catch
            {
                MessageBox.Show("Ha Sucedido un Error Intente de Nuevo");
            }
        }

        public static void Merge(
            string[] caminhosArquivosOrigem,
            string caminhoNovoArquivoPDF)
        {
            using (FileStream stream =
                new FileStream(caminhoNovoArquivoPDF, FileMode.Create))
            {
                iTextSharp.text.Document documento = new iTextSharp.text.Document();
                PdfCopy pdfCopy = new PdfCopy(documento, stream);
                documento.Open();

                foreach (string arquivoOrigem in caminhosArquivosOrigem)
                {
                    pdfCopy.AddDocument(new PdfReader(arquivoOrigem));
                }

                if (documento != null)
                    documento.Close();
            }
        }

        private void btnSeleccionar1_Click(object sender, EventArgs e)
        {
            this.openFileDialog3.Filter =
           "Images (*.PDF)|*.PDF|" +
           "All files (*.*)|*.*";
            this.openFileDialog3.Title = "Seleccionar Archivo Pdf";

            if (openFileDialog3.ShowDialog() == DialogResult.OK)
                lstArquivosOrigem.Items.AddRange(openFileDialog3.FileNames);
        }


        private void btnSeleccionar1_MouseHover(object sender, EventArgs e)
        {
            btnSeleccionar1.BackColor = System.Drawing.Color.White;
            btnSeleccionar1.ForeColor = System.Drawing.Color.Black;
        }

        private void btnSeleccionar1_MouseLeave(object sender, EventArgs e)
        {
            btnSeleccionar1.BackColor = System.Drawing.Color.Green;
            btnSeleccionar1.ForeColor = System.Drawing.Color.White;
        }

        private void btnUnirPdf_MouseHover(object sender, EventArgs e)
        {
            btnUnirPdf.BackColor = System.Drawing.Color.Red;
            btnUnirPdf.ForeColor = System.Drawing.Color.White;
        }

        private void btnUnirPdf_MouseLeave(object sender, EventArgs e)
        {
            btnUnirPdf.BackColor = System.Drawing.Color.Orange;
            btnUnirPdf.ForeColor = System.Drawing.Color.Black;
        }

        /**************************************Unir PDF*****************************/

        public class MessageBoxTemporal
        {
            System.Threading.Timer IntervaloTiempo;
            string TituloMessageBox;
            string TextoMessageBox;
            int TiempoMaximo;
            IntPtr hndLabel = IntPtr.Zero;
            bool MostrarContador;

            MessageBoxTemporal(string texto, string titulo, int tiempo, bool contador)
            {
                TituloMessageBox = titulo;
                TiempoMaximo = tiempo;
                TextoMessageBox = texto;
                MostrarContador = contador;
              

                if (TiempoMaximo > 99) return; //Máximo 99 segundos
                IntervaloTiempo = new System.Threading.Timer(EjecutaCada1Segundo,
                    null, 1000, 1000);
                if (contador)
                {
                    DialogResult ResultadoMensaje = MessageBox.Show(texto + "\r\nEste mensaje se cerrará dentro de " +
                        TiempoMaximo.ToString("00") + " segundos ...", titulo);
                    if (ResultadoMensaje == DialogResult.OK) IntervaloTiempo.Dispose();
                }
                else
                {
                    DialogResult ResultadoMensaje = MessageBox.Show(texto + "...", titulo);
                    if (ResultadoMensaje == DialogResult.OK) IntervaloTiempo.Dispose();
                }
            }
            public static void Show(string texto, string titulo, int tiempo, bool contador)
            {
                new MessageBoxTemporal(texto, titulo, tiempo, contador);
            }
            void EjecutaCada1Segundo(object state)
            {
                TiempoMaximo--;
                if (TiempoMaximo <= 0)
                {
                    IntPtr hndMBox = FindWindow(null, TituloMessageBox);
                    if (hndMBox != IntPtr.Zero)
                    {
                        SendMessage(hndMBox, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                        IntervaloTiempo.Dispose();
                    }
                }
                else if (MostrarContador)
                {
                    // Ha pasado un intervalo de 1 seg:
                    if (hndLabel != IntPtr.Zero)
                    {
                        SetWindowText(hndLabel, TextoMessageBox +
                            "\r\nEste mensaje se cerrará dentro de " +
                            TiempoMaximo.ToString("00") + " segundos");
                    }
                    else
                    {
                        IntPtr hndMBox = FindWindow(null, TituloMessageBox);
                        if (hndMBox != IntPtr.Zero)
                        {
                            // Ha encontrado el MessageBox, busca ahora el texto
                            hndLabel = FindWindowEx(hndMBox, IntPtr.Zero, "Static", null);
                            if (hndLabel != IntPtr.Zero)
                            {
                                // Ha encontrado el texto porque el MessageBox
                                // solo tiene un control "Static".
                                SetWindowText(hndLabel, TextoMessageBox +
                                    "\r\nEste mensaje se cerrará dentro de " +
                                    TiempoMaximo.ToString("00") + " segundos");
                            }
                        }
                    }
                }
            }
            const int WM_CLOSE = 0x0010;
            [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
            static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
            [System.Runtime.InteropServices.DllImport("user32.dll",
                CharSet = System.Runtime.InteropServices.CharSet.Auto)]
            static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
            [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true,
                CharSet = System.Runtime.InteropServices.CharSet.Auto)]
            static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter,
                string lpszClass, string lpszWindow);
            [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true,
                CharSet = System.Runtime.InteropServices.CharSet.Auto)]
            static extern bool SetWindowText(IntPtr hwnd, string lpString);
        }


        /**********************************Dividir***************/
        private void btnSelecDividir_Click(object sender, EventArgs e)
        {
            String file_name = string.Empty;
            this.openFileDialog4.Filter =
           "Documentos(*.PDF;)| *.PDF; | " +
           "All files (*.*)|*.*";
            this.openFileDialog4.Title = "Seleccionar Documento";

            if (openFileDialog4.ShowDialog() == DialogResult.OK)
            {
                dir = openFileDialog4.FileName;
                destino = Path.GetFileName(dir);
                txtDividir.Text = destino;
            }
        }

        private void btnDividir_Click(object sender, EventArgs e)
        {

            try
            {
                MessageBoxTemporal.Show("Espere mientras se Convierte la Imagen a PDF", " Baymax v1.5", 3, false);
                FileInfo file = new FileInfo(dir);
                string name = file.Name.Substring(0, file.Name.LastIndexOf("."));

                using (PdfReader reader = new PdfReader(dir))
                {

                    for (int pagenumber = 1; pagenumber <= reader.NumberOfPages; pagenumber++)
                    {
                        string filename = pagenumber.ToString() + ".pdf";

                        iTextSharp.text.Document document = new iTextSharp.text.Document();
                        PdfCopy copy = new PdfCopy(document, new FileStream(@salidaDocumentos+"\\" + Path.GetFileName(dir) + "-" + filename, FileMode.Create));

                        document.Open();

                        copy.AddPage(copy.GetImportedPage(reader, pagenumber));

                        document.Close();
                    }

                }
            }
            catch { MessageBox.Show("Ha Sucedido un Error Intente de Nuevo"); }

            MessageBox.Show("Se ha Dividio el Pdf Correctamente", "  Baymax v1.5", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            txtDividir.Clear();
        }

        private void btnSelecDividir_MouseHover(object sender, EventArgs e)
        {
            btnSelecDividir.BackColor = System.Drawing.Color.White;
            btnSelecDividir.ForeColor = System.Drawing.Color.Black;
        }

        private void btnSelecDividir_MouseLeave(object sender, EventArgs e)
        {
            btnSelecDividir.BackColor = System.Drawing.Color.Green;
            btnSelecDividir.ForeColor = System.Drawing.Color.White;
        }


        private void btnDividirPdf_MouseHover(object sender, EventArgs e)
        {
            btnDividir.BackColor = System.Drawing.Color.Red;
            btnDividir.ForeColor = System.Drawing.Color.White;
        }

        private void btnDividirPdf_MouseLeave(object sender, EventArgs e)
        {
            btnDividir.BackColor = System.Drawing.Color.Orange;
            btnDividir.ForeColor = System.Drawing.Color.Black;
        }


        private void btnSelecDesproteger_Click(object sender, EventArgs e)
        {
            this.openFileDialog6.Filter =
          "Images (*.PDF)|*.PDF|" +
          "All files (*.*)|*.*";
            this.openFileDialog6.Title = "Seleccionar Archivo Pdf";

            if (openFileDialog6.ShowDialog() == DialogResult.OK)
                lsDesproteger.Items.AddRange(openFileDialog6.FileNames);

        }

        private void btnDesproteger_Click(object sender, EventArgs e)
        {
            try
            {
               
                    foreach (string arquivoOrigem1 in lsDesproteger.Items.Cast<string>().ToArray())
                    {
                        var api = new LovePdfApi("project_public_681de94b8592545e00e8ea3aa6b85cef_zjS6D000863c42d932d8966da0023c96b8731", "secret_key_f9b5c586555f47c6b0a1c542db712ba9_8qTgT48d1dd3f7a50c806bcd52fe11aa51328");

                        //create compress task
                        var task = api.CreateTask<UnlockTask>();
                        MessageBoxTemporal.Show("Espere mientras se Desprotege el PDF", " Baymax v1.5", 3, false);
                    //file variable contains server file name

                    var file = task.AddFile(arquivoOrigem1, task.TaskId, "test");

                    //proces added files
                    //time var will contains information about time spent in process
                    var time = task.Process();

                    //download files to specific folder
                    task.DownloadFile(@salidaDocumentos+"\\");
                    }
                    MessageBox.Show("Se ha Desprotegido el Documento/s\nBusque su Archivo", " Baymax v1.5", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    lsDesproteger.Items.Clear();
                }

               
            
            catch
            {
              
                MessageBox.Show("Ha Sucedido un Error Intente de Nuevo"); }
        }

       

        private void ch0_Click(object sender, EventArgs e)
        {
            ch0.Checked = true;
            ch90.Checked = false;
            ch180.Checked = false;
            ch270.Checked = false;
        }

        private void ch90_Click(object sender, EventArgs e)
        {
            ch0.Checked = false;
            ch90.Checked = true;
            ch180.Checked = false;
            ch270.Checked = false;
        }

        private void ch180_Click(object sender, EventArgs e)
        {
            ch0.Checked = false;
            ch90.Checked = false;
            ch180.Checked = true;
            ch270.Checked = false;
        }

        private void ch270_Click(object sender, EventArgs e)
        {
            ch0.Checked = false;
            ch90.Checked = false;
            ch180.Checked = false;
            ch270.Checked = true;
        }

        private void btnSeleccionarDirectorio_Click(object sender, EventArgs e)
        {
            txtDirectorio.Clear();
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string CurrentDirectory;
               
                CurrentDirectory = Path.GetFullPath(folderBrowserDialog1.SelectedPath);
               
                if(CurrentDirectory==null)
                {
                    CurrentDirectory = Path.GetPathRoot(folderBrowserDialog1.SelectedPath);
                }
              
                    txtDirectorio.Text = CurrentDirectory.ToString();
                    salidaDocumentos= CurrentDirectory.ToString();

                if (string.IsNullOrEmpty(salidaDocumentos))
                {
                    tabControl1.Enabled = false;
                }
                else
                {
                    tabControl1.Enabled = true;
                }
            }
        }

    

        private void btnDirectorioAbrir_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@salidaDocumentos+"\\");
        }
    }



}
