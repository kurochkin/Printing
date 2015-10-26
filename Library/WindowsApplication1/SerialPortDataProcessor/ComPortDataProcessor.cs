using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;


namespace SerialPortDataProcessor
{
    [ComVisible(true)]
    [ProgId("SerialPortDataProcessor.ComPortDataProcessor")]
    [Guid("BA66896F-B746-4F87-B45E-AAD7CF11C6B9")]
    public class ComPortDataProcessor
    {


        private SerialPort _serialPort;
        private  double _lastValue;

        private Thread _workingThread;
        private bool _isProcessing = false;

        public void InitReading(string portName, int dataBits, int baudRate)
        {
            _serialPort = new SerialPort(portName);

            _serialPort.BaudRate = baudRate;//9600;
            _serialPort.Parity = Parity.None;
            _serialPort.StopBits = StopBits.One;
            _serialPort.DataBits = dataBits;//8;
            _serialPort.Handshake = Handshake.None;

        }


        public void ThreadRoutine(object o)
        {

            if (!_serialPort.IsOpen)
                _serialPort.Open();

            try
            {

                while (_isProcessing)
                {
                    _serialPort.ReadTimeout = 200;
                    try
                    {
                        string readedLine = _serialPort.ReadLine();
                        _lastValue = ParseReceivedString(readedLine);
                    }
                    catch (TimeoutException timeoutException) { }
                }
            }
            finally
            {
                _serialPort.Close();
                //_readedValues.Clear();
            }
        }

        public void StartReading()
        {
            if (_isProcessing == false)
            {
                _isProcessing = true;
                _workingThread = new Thread(ThreadRoutine);
                _workingThread.Start();
            }
        }

        public void StopReading()
        {
            _isProcessing = false;
        }

        public string ReadedData()
        {
            string resString = _lastValue.ToString();
            return resString;
        }


        private double ParseReceivedString(string reseivedString)
        {
            bool isStartNumber = false;
            string allNumbersStrin = String.Empty;
            for (int i = 0; i < reseivedString.Length; i++)
            {
                var currentChar = reseivedString[i];
                if (char.IsNumber(currentChar) || currentChar == '.')
                {
                    allNumbersStrin += currentChar;
                    isStartNumber = true;
                }
                else
                {
                    if (isStartNumber)
                    {
                        allNumbersStrin += ';'; //separator
                        isStartNumber = false;
                    }
                }
            }

            var numbersArr = allNumbersStrin.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

            List<double> res = new List<double>();
            foreach (string numStr in numbersArr)
            {
                double result;
                double.TryParse(numStr, NumberStyles.Any, CultureInfo.InvariantCulture, out result);
                res.Add(result);
            }


            return res[0];
        }

      

    }
}
