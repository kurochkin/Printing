using System;
using System.Data;

namespace SerialPortDataProcessor.Printing.Data
{
    public class LabelItem
    {
        public LabelItem(IDataRecord dataRow)
        {
            var labelTypeStr = dataRow["LabelType"].ToString();
            LabelType = (LabelTypesEnum)Enum.Parse(typeof(LabelTypesEnum), labelTypeStr);
            StartX = (int)dataRow["StartX"];
            LabelText = (string)dataRow["LabelText"];
            StartY = (int)dataRow["StartY"];
            FontSize = (int)dataRow["FontSize"];
            IsBold = (bool)dataRow["IsBold"];
            PageNum = (int)dataRow["PageNum"];
            FontName = (string)dataRow["FontName"];
            IsItalic = (bool)dataRow["IsItalic"];
        }

        public LabelTypesEnum LabelType { get; set; }
        public string LabelText { get; set; }
        public int StartX { get; set; }
        public int StartY { get; set; }
        public int FontSize { get; set; }
        public bool IsBold { get; set; }
        public int PageNum { get; set; }
        public string FontName { get; set; }
        public bool IsItalic { get; set; }



    }
}
