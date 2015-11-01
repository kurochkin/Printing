using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SerialPortDataProcessor;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private SerialPortDataProcessor.ComPortDataProcessor _portDataProcessor;
        private SerialPortDataProcessor.PrintLabelProcessor _printLabelProcessor = new PrintLabelProcessor();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _portDataProcessor.StartReading();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _portDataProcessor = new ComPortDataProcessor();
            _portDataProcessor.InitReading("COM1", 8, 9600);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _portDataProcessor.StopReading();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show(_portDataProcessor.ReadedData());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            _printLabelProcessor.Init();
            _printLabelProcessor.PrintLabelV2(100);
        }
    }
}
