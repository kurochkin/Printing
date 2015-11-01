using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;

using System.Windows.Forms;
using WindowsFormsApplication1.Data;
using ZXing;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private List<Bitmap> _frontBitmaps;
        private List<Bitmap> _backBitmaps;
        private List<Bitmap> _v2Bitmaps;

        //private Bitmap _commonBitmap;


        private PrintLabelSettings GetPrintLabelSettings()
        {
            using (var conn = new SqlConnection(txtConnectionString.Text))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = "usp_GetPrintLabelSettings";
                    var reader = cmd.ExecuteReader();

                    foreach (IDataRecord dataRecord in reader)
                        return new PrintLabelSettings(dataRecord);
                }
            }
            return null;
        }

        private List<Bitmap> ProcessLabel(string spName, bool mirrorTransfom, int width, int height)
        {
            List<LabelItem> labelItems = new List<LabelItem>();

            int currentPange = 1;

            List<Bitmap> bmpLst = new List<Bitmap>();
            bmpLst.Add(new Bitmap(width, height));

            using (var conn = new SqlConnection(txtConnectionString.Text))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = spName;
                    cmd.Parameters.Add("Id", txtId.Text);
                    var reader = cmd.ExecuteReader();
                    foreach (IDataRecord daraRecord in reader)
                        labelItems.Add(new LabelItem(daraRecord));
                }
            }

            var settings = GetPrintLabelSettings();

            foreach (var labelItem in labelItems)
            {
                if (currentPange != labelItem.PageNum)
                {
                    bmpLst.Add(new Bitmap(width, height));
                    currentPange++;
                }

                using (var g = Graphics.FromImage(bmpLst[currentPange - 1]))
                {
                    if (mirrorTransfom)
                    {
                        g.TranslateTransform(width, height);
                        g.ScaleTransform(-1, -1);
                    }

                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;


                    ProcessLabelItem(labelItem, g, settings);
                }
            }
            return bmpLst;
        }

        private static void ProcessLabelItem(LabelItem labelItem, Graphics g, PrintLabelSettings settings)
        {
            switch (labelItem.LabelType)
            {
                case LabelTypesEnum.Label:
                    g.DrawString(labelItem.LabelText,
                        new Font("Arial", labelItem.FontSize, labelItem.IsBold ? FontStyle.Bold : FontStyle.Regular),
                        Brushes.Black, labelItem.StartX, labelItem.StartY);
                    break;
                case LabelTypesEnum.BarCode:
                    var content = labelItem.LabelText;

                    var writer = new BarcodeWriter
                    {
                        Format = BarcodeFormat.CODE_128,
                        Options = new ZXing.QrCode.QrCodeEncodingOptions
                        {
                            ErrorCorrection = ZXing.QrCode.Internal.ErrorCorrectionLevel.H,
                            Width = settings.BarCodeMaxWidth,
                            Height = settings.BarCodeHeight,
                            PureBarcode = true,
                        }
                    };
                    var barCodeBmp = writer.Write(content);
                    g.DrawImageUnscaled(barCodeBmp, labelItem.StartX, labelItem.StartY);
                    break;
                case LabelTypesEnum.Stamp:
                    var pen = new Pen(Color.Black, 2);
                    g.DrawEllipse(pen, labelItem.StartX, labelItem.StartY, settings.StampDiameter, settings.StampDiameter);
                    g.DrawString(labelItem.LabelText,
                        new Font("Arial", labelItem.FontSize, labelItem.IsBold ? FontStyle.Bold : FontStyle.Regular),
                        Brushes.Black, labelItem.StartX + 2, labelItem.StartY + 11);
                    break;
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                _frontBitmaps = ProcessLabel("usp_PrintLabelFront", false, 220, 365);
                pictureBox1.Image = _frontBitmaps[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                _backBitmaps = ProcessLabel("usp_PrintLabelBack", true, 220, 365);
                pictureBox2.Image = _backBitmaps[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void PrintImage(Image img)
        {
            PrintImage(img, true);
        }

        Image bmIm;

        private void PrintImage(Image img, bool showPreview)
        {
            PrintImage(img, showPreview, null);
        }
        private void PrintImage(Image img, bool showPreview, string pageInfo)
        {
            try
            {
                var settings = GetPrintLabelSettings();

                bmIm = img;

                if (string.IsNullOrEmpty(pageInfo) == false)
                    DrawPageInfo(pageInfo, settings);

                PrintDocument pd = new PrintDocument();
                pd.PrinterSettings.PrinterName = "Microsoft XPS Document Writer";
                pd.OriginAtMargins = true;
                pd.DefaultPageSettings.Margins = new Margins(settings.MarginLeft, 0, settings.MarginTop, 0);

                //pd.OriginAtMargins = false;
                pd.PrintPage += pd_PrintPage;
                pd.DefaultPageSettings.Landscape = false;
                //pd.Print();
                if (showPreview)
                {
                    PrintPreviewDialog dialog = new PrintPreviewDialog();
                    dialog.UseAntiAlias = true;

                    dialog.Document = pd;
                    dialog.ShowDialog();
                }
                else
                {
                    pd.Print();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DrawPageInfo(string pageInfo, PrintLabelSettings settings)
        {
            using (var g = Graphics.FromImage(bmIm))
            {
                //g.SmoothingMode = SmoothingMode.HighQuality;
                g.TextRenderingHint = TextRenderingHint.SystemDefault;
                //g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                //g.PixelOffsetMode = PixelOffsetMode.HighQuality;


                g.DrawString(pageInfo, new Font("Arial", 14, FontStyle.Bold), Brushes.Black, settings.PageCounterX, settings.PageCounterY);
            }
        }

        void pd_PrintPage(object sender, PrintPageEventArgs e)
        {
            e.Graphics.PageUnit = GraphicsUnit.Pixel;
            //e.Graphics.CompositingMode = CompositingMode.SourceCopy;
            //e.Graphics.CompositingQuality = CompositingQuality.HighQuality;


            //e.Graphics.SmoothingMode = SmoothingMode.HighQuality;
            e.Graphics.TextRenderingHint = TextRenderingHint.SystemDefault;
            //e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            //e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

            e.Graphics.DrawImage(bmIm, 0, 0, bmIm.Width, bmIm.Height);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PrintImage(_frontBitmaps[0]);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PrintImage(_backBitmaps[0]);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            button1_Click(sender, e);
            button3_Click(sender, e);

            int pagesCount = _backBitmaps.Count;
            int currentPage = 0;

            foreach (var bitmap in _backBitmaps)
            {
                currentPage++;
                PrintImage(bitmap, false);
                PrintImage(_frontBitmaps[0], false, pagesCount > 1 ? String.Format("{0}\\{1}", currentPage, pagesCount) : String.Empty);
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                _v2Bitmaps = ProcessLabel("usp_PrintLabelV2", false, 220, 365);
                pictureBox3.Image = _v2Bitmaps[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (_v2Bitmaps == null || _v2Bitmaps.Count == 0)
                button6_Click(sender, e);

            PrintImage(_v2Bitmaps[0]);
        }
    }
}
