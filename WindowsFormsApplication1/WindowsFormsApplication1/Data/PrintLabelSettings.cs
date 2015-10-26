using System.Data;

namespace WindowsFormsApplication1.Data
{
    public class PrintLabelSettings
    {
        public int MarginLeft { get; set; }
        public int MarginTop { get; set; }
        public int BarCodeHeight { get; set; }
        public int BarCodeMaxWidth { get; set; }
        public int ImagesMargin { get; set; }
        public int StampDiameter { get; set; }
        public int PageCounterX { get; set; }
        public int PageCounterY { get; set; }


        public PrintLabelSettings(IDataRecord dataRecord)
        {
            MarginLeft = (int) dataRecord["MarginLeft"];
            MarginTop = (int) dataRecord["MarginTop"];
            BarCodeHeight = (int) dataRecord["BarCodeHeight"];
            BarCodeMaxWidth = (int) dataRecord["BarCodeMaxWidth"];
            ImagesMargin = (int)dataRecord["ImagesMargin"];
            StampDiameter = (int)dataRecord["StampDiameter"];
            PageCounterX = (int)dataRecord["PageCounterX"];
            PageCounterY = (int)dataRecord["PageCounterY"];
        }
    }
}
