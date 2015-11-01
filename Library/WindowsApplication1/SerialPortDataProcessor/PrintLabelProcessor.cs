using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Globalization;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;
using SerialPortDataProcessor.Printing.Data;
using ZXing;


namespace SerialPortDataProcessor
{
    [ComVisible(true)]
    [Guid("068CC3EF-18E9-473B-B711-094790B1CE27")]
    [ProgId("SerialPortDataProcessor.PrintLabelProcessor")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class PrintLabelProcessor
    {
        private string _connectionString;
        private int _id;
        public void Init()
        {
            _connectionString = Configuration.GetConnectionString();
        }

        public void PrintLabel(int id)
        {
            _id = id;
            //_frontBitmaps = ProcessLabel("usp_PrintLabelFront", false, 220, 365);
            _backBitmaps = ProcessLabel("usp_PrintLabelBack", true, 220, 365);

            int pagesCount = _backBitmaps.Count;
            int currentPage = 0;

            foreach (var bitmap in _backBitmaps)
            {
                currentPage++;
                PrintImage(bitmap, false);
                _frontBitmaps = ProcessLabel("usp_PrintLabelFront", false, 220, 365);
                PrintImage(_frontBitmaps[0], false, pagesCount > 1 ? String.Format("{0} \\ {1}", currentPage, pagesCount) : String.Empty);
            }
        }

        public void PrintLabelV2(int id)
        {
            _id = id;
            _v2Bitmaps = ProcessLabel("usp_PrintLabelV2", false, 220, 365);

            if (_v2Bitmaps.Count == 0) return;

            foreach (var bitmap in _v2Bitmaps)
                PrintImage(bitmap, false, null);
        }

        private List<Bitmap> _frontBitmaps;
        private List<Bitmap> _backBitmaps;
        private List<Bitmap> _v2Bitmaps;



        private PrintLabelSettings GetPrintLabelSettings()
        {
            using (var conn = new SqlConnection(_connectionString))
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

            int currentPage = 1;

            List<Bitmap> bmpLst = new List<Bitmap>();
            bmpLst.Add(new Bitmap(width, height));

            using (var conn = new SqlConnection(_connectionString))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = spName;
                    cmd.Parameters.Add("Id", _id);
                    var reader = cmd.ExecuteReader();
                    foreach (IDataRecord daraRecord in reader)
                        labelItems.Add(new LabelItem(daraRecord));
                }
            }

            var settings = GetPrintLabelSettings();

            foreach (var labelItem in labelItems)
            {
                if (currentPage != labelItem.PageNum)
                {
                    bmpLst.Add(new Bitmap(width, height));
                    currentPage++;
                }

                using (var g = Graphics.FromImage(bmpLst[currentPage - 1]))
                {
                    if (mirrorTransfom)
                    {
                        g.TranslateTransform(width, height);
                        g.ScaleTransform(-1, -1);
                    }

                    //g.SmoothingMode = SmoothingMode.HighQuality;
                    g.TextRenderingHint = TextRenderingHint.SystemDefault;
                    //g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    //g.PixelOffsetMode = PixelOffsetMode.HighQuality;


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

        //private void PrintImage(Image img)
        //{
        //    PrintImage(img, true);
        //}

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
                pd.PrinterSettings.PrinterName = Configuration.GetPrinterName();
                pd.OriginAtMargins = true;
                pd.DefaultPageSettings.Margins = new Margins(settings.MarginLeft, 0, settings.MarginTop, 0);
                //pd.OriginAtMargins = false;
                pd.PrintPage += pd_PrintPage;
                pd.DefaultPageSettings.Landscape = false;
                pd.Print();
                //if (showPreview)
                //{
                //    PrintPreviewDialog dialog = new PrintPreviewDialog();
                //    dialog.UseAntiAlias = true;

                //    dialog.Document = pd;
                //    dialog.ShowDialog();
                //}
                //else
                //{
                //    pd.Print();
                //}
            }
            catch (Exception ex)
            {
                throw;
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

                g.DrawString(pageInfo, new Font("Arial", settings.PageCounterFontSize, settings.PageCounterIsBold ? FontStyle.Bold : FontStyle.Regular), Brushes.Black, settings.PageCounterX, settings.PageCounterY);
            }
        }

        void pd_PrintPage(object sender, PrintPageEventArgs e)
        {
            e.Graphics.PageUnit = GraphicsUnit.Pixel;
            e.Graphics.DrawImage(bmIm, 0, 0, bmIm.Width, bmIm.Height);
        }

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    PrintImage(_frontBitmaps[0]);
        //}

        //private void button4_Click(object sender, EventArgs e)
        //{
        //    PrintImage(_backBitmaps[0]);
        //}



    }
}
