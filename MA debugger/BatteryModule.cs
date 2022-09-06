using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Timers;
using Timer = System.Threading.Timer;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;

namespace MA_debugger
{
    class BatteryModule
    {
        private int _mainLoopPeriod;
        private int _pollingPeriod;
        private readonly MemoryBms _memory = new MemoryBms();
        private readonly Form1 _form;
        private readonly SerialPort _serialPort;
        private readonly Timer _timerMainLoop;
        private readonly Timer _timerLoopGetMeasurements;
        private bool _isPollEnable = false;
        private UartCommand _lastCommand = UartCommand.Null;
        private readonly byte[] _receiveBufferStatusRegisters = new byte[Marshal.SizeOf(typeof(StatusRegisters))];
        private int _countOfReceiveBuffer = 0;
        private readonly byte[] _bufferSettingsCharge = new byte[Marshal.SizeOf(typeof(SettingsCharge))];
        private readonly byte[] _bufferSettingsAlarm = new byte[Marshal.SizeOf(typeof(SettingsAlarm))];
        private readonly Queue<UartCommand> _unsentСommands = new Queue<UartCommand>();
        private ushort _calibrationValue = 0;
        private uint _calibrationCoefficient;
        private int _step = 0;
        private const int ChecksumConstant = 44111;
        private readonly System.Timers.Timer _uartRxErrTimer;
        private Excel.Application _excelApp;
        private Excel.Workbooks _excelAppWorkbooks;
        private Excel.Workbook _excelAppWorkbook;
        private Excel.Sheets _excelSheets;
        private Excel.Worksheet _excelWorkSheet;
        private int _excelMeasurementsPeriod = 60; // Количество измерений для выведения среднего для записи в Excel
        private const int IndexOfFirstColumnVoltage = 2;
        private const int IndexOfFirstLineVoltage = 4;
        private const int IndexColumnOfData = 4;
        private const int IndexLineOfData = 2;
        private const int IndexColumnOfTime = 1;
        private const int IndexLineOfDuration = 2;
        private const int IndexColumnOfDuration = 11;
        private const int IndexLineOfStartVoltage = 3;
        private const int IndexColumnOfStartVoltage = 11;
        private const int IndexLineOfEndVoltage = 4;
        private const int IndexColumnOfEndVoltage = 11;
        private int _currentLine = IndexOfFirstLineVoltage;
        private const int NumberOfCells = 4;
        private readonly int[] _cellVoltage = new int[NumberOfCells];
        private int _indexOfMeasurement = 0;
        private bool _isFirstDataAddToExcel = false;

        internal enum UartCommand : byte
        { // Commands: 0x80 - 0xEF
            Null = 0x00,
            UartCommandPcEnableDebugging = 0x80, /*!< Command from PC. Enable debugging discrete devices. */
            UartCommandPcDisableDebugging = 0x81, /*!< Command from PC. Disable debugging discrete devices. */
            UartCommandPcBalancingCell1Enable = 0x82, /*!< Command from PC. Enable balancing cell 1. */
            UartCommandPcBalancingCell2Enable = 0x83, /*!< Command from PC. Enable balancing cell 2. */
            UartCommandPcBalancingCell3Enable = 0x84, /*!< Command from PC. Enable balancing cell 3. */
            UartCommandPcBalancingCell4Enable = 0x85, /*!< Command from PC. Enable balancing cell 4. */
            UartCommandPcBalancingCell1Disable = 0x86, /*!< Command from PC. Disable balancing cell 1. */
            UartCommandPcBalancingCell2Disable = 0x87, /*!< Command from PC. Disable balancing cell 2. */
            UartCommandPcBalancingCell3Disable = 0x88, /*!< Command from PC. Disable balancing cell 3. */
            UartCommandPcBalancingCell4Disable = 0x89, /*!< Command from PC. Disable balancing cell 4. */
            UartCommandPcContactorEnable = 0x8A, /*!< Command from PC. Enable contactor. */
            UartCommandPcContactorDisable = 0x8B, /*!< Command from PC. Disable contactor. */
            UartCommandPcChargerEnable = 0x8C, /*!< Command from PC. Enable charger. */
            UartCommandPcChargerDisable = 0x8D, /*!< Command from PC. Disable charger. */
            UartCommandPcMotorEnable = 0x8E, /*!< Command from PC. Enable motor. */
            UartCommandPcMotorDisable = 0x8F, /*!< Command from PC. Disable motor. */

            UartCommandPcCalibrationU1 = 0xE1, /*!< Command from PC. Calibration U1. */
            UartCommandPcCalibrationU2 = 0xE2, /*!< Command from PC. Calibration U2. */
            UartCommandPcCalibrationU3 = 0xE3, /*!< Command from PC. Calibration U3. */
            UartCommandPcCalibrationU4 = 0xE4, /*!< Command from PC. Calibration U4. */
            UartCommandPcCalibrationTemp = 0xE5, /*!< Command from PC. Calibration temperature. */
            UartCommandPcSendStatusRegisters = 0xE6, /*!< Command from PC. Send the status registers struct. */
            UartCommandPcSendSettingsCharge = 0xE7, /*!< Command from PC. Send settings charge. */
            UartCommandPcSendSettingsAlarm = 0xE8, /*!< Command from PC. Send settings alarm. */
            UartCommandPcUpdateSettingsCharge = 0xE9, /*!< Command from PC. Update settings BMS. */
            UartCommandPcUpdateSettingsAlarm = 0xEA, /*!< Command from PC. Update settings alarm. */
            UartCommandPcSaveSettings = 0xEB, /*!< Command from PC. Save settings to flash. */
            UartCommandPcDefaultCalibration = 0xEC, /*!< Command from PC. Set default calibration settings. */
            UartCommandPcDefaultCharge = 0xED, /*!< Command from PC. Set default charge settings. */
            UartCommandPcDefaultAlarm = 0xEE, /*!< Command from PC. Set default alarm settings. */

            UartCommandOk = 0xFA
        }

        public BatteryModule(Form1 form, SerialPort serialPort, int pollingPeriod)
        {
            _form = form;
            _serialPort = serialPort;
            _pollingPeriod = pollingPeriod;
            _mainLoopPeriod = _pollingPeriod / 100;
            _timerLoopGetMeasurements = new Timer(LoopGetMeasurements, null, 0, _pollingPeriod);
            _timerMainLoop = new Timer(MainLoop, null, 0, _mainLoopPeriod);
            _serialPort.DataReceived += ComPortDataReceived;
            
            // Initialise UartRxErrTimer:
            _uartRxErrTimer = new System.Timers.Timer(100); // Message received time, ms
            _uartRxErrTimer.Elapsed += TimerEventUartRxErrProcessor;
            _uartRxErrTimer.AutoReset = true;
        }

        public void ExcelOpenBook()
        {
            try
            {// Присоединение к открытому приложению Excel (если оно открыто), имхо так тру, ибо 2 excel процесса в памяти не кошерно
                _excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                _excelApp = new Excel.Application(); // Если нет открытого, то создаём новое приложение
            }
            try
            {
                _excelApp.Visible = true;
                _excelApp.DisplayAlerts = false;
                _excelAppWorkbooks = _excelApp.Workbooks; // Получаем список открытых книг
                _excelAppWorkbook = GetWorkbook(_excelAppWorkbooks);
                _excelSheets = _excelAppWorkbook.Worksheets; // Получаем список листов в нашей книге
                _excelWorkSheet = _excelSheets[1]; // Берем первый лист
                _excelWorkSheet.Copy(Type.Missing, _excelSheets[_excelSheets.Count]); // Копируем первый лист в конец
                _excelWorkSheet = _excelSheets[_excelSheets.Count];
                int numberOfLastScheet = Int32.Parse(_excelSheets[_excelSheets.Count - 1].Name) + 1; // Номер следующего листа
                _excelWorkSheet.Name = numberOfLastScheet.ToString();
            }
            catch (Exception theException)
            {
                var errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, @"Error");
            }
        }

        private Excel.Workbook GetWorkbook(Excel.Workbooks excelAppWorkbooks)
        {
            Excel.Workbook excelAppWorkbook = null;
            var fileName = @"D:\_VS_PROJECTS\BMS4S debugger\MA debugger\bin\Debug\Отчет.xlsx";
            bool isBookOpened = false;
            for (int i = 1; i <= excelAppWorkbooks.Count; i++)
            {
                if (excelAppWorkbooks[i].FullName == fileName)
                {
                    excelAppWorkbook = excelAppWorkbooks[i]; // Устанавливаем ссылку на нашу книгу
                    isBookOpened = true;
                }
            }
            if (!isBookOpened)
            {
                excelAppWorkbook = _excelApp.Workbooks.Open(fileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
            }
            return excelAppWorkbook;
        }

        public void ExcelSaveBook()
        {
            _excelAppWorkbook.Save();
        }

        public void ResetMeasurements()
        {
            for (int i = 0; i < _cellVoltage.Length; i++)
            {
                _cellVoltage[i] = 0;
            }
            _indexOfMeasurement = 0;
        }

        private void UpdateMeasurementsToExcel()
        {
            if (_indexOfMeasurement < _excelMeasurementsPeriod)
            {
                _cellVoltage[0] += _memory.StatusRegisters.cellVoltage1;
                _cellVoltage[1] += _memory.StatusRegisters.cellVoltage2;
                _cellVoltage[2] += _memory.StatusRegisters.cellVoltage3;
                _cellVoltage[3] += _memory.StatusRegisters.cellVoltage4;
                _indexOfMeasurement++;
            } else
            {
                for (int i = 0; i < _cellVoltage.Length; i++)
                {
                    _cellVoltage[i] /= _excelMeasurementsPeriod;
                }
                AddDataToExcel();
                ResetMeasurements();
            }
            if (!_isFirstDataAddToExcel)
            {
                AddStartDataToExcel();
                _isFirstDataAddToExcel = true;
            }
        }

        public bool IsAddStartDataToExcel
        {
            set => _isFirstDataAddToExcel = value;
        }

        private void AddStartDataToExcel()
        {
            _excelWorkSheet.Cells[IndexLineOfData, IndexColumnOfData] = DateTime.Now;
            _currentLine = IndexOfFirstLineVoltage;
            AddCellsVoltagesToExcel();
            int batteryVoltage = _memory.StatusRegisters.batteryVoltage;
            _excelWorkSheet.Cells[IndexLineOfStartVoltage, IndexColumnOfStartVoltage] = ConvertMilliVoltToVolt(batteryVoltage);
        }

        private void AddDataToExcel()
        {
            AddCellsVoltagesToExcel();
            int batteryVoltage = _memory.StatusRegisters.batteryVoltage;
            _excelWorkSheet.Cells[IndexLineOfEndVoltage, IndexColumnOfEndVoltage] = ConvertMilliVoltToVolt(batteryVoltage);
        }

        private void AddCellsVoltagesToExcel()
        {
            AddCurrentTimeToExcel(_currentLine);
            for (int i = IndexOfFirstColumnVoltage; i < _cellVoltage.Length + IndexOfFirstColumnVoltage; i++)
            {
                _excelWorkSheet.Cells[_currentLine, i] = ConvertMilliVoltToVolt(_cellVoltage[i - IndexOfFirstColumnVoltage]);
            }
            Excel.ChartObject chartTemplate = (Excel.ChartObject)_excelWorkSheet.ChartObjects("chart 1");
            var dataRange = _excelWorkSheet.Range["A3", "E" + _currentLine];
            chartTemplate.Chart.SetSourceData(dataRange, Excel.XlRowCol.xlColumns);
            _currentLine++;
        }

        private float ConvertMilliVoltToVolt(int milliVolt)
        {
            return (float)milliVolt / 1000;
        }

        private void AddCurrentTimeToExcel(int line)
        {
            _excelWorkSheet.Range[_excelWorkSheet.Cells[line, IndexColumnOfTime], _excelWorkSheet.Cells[line, IndexColumnOfTime]].NumberFormat = "ч:мм:сс";
            _excelWorkSheet.Cells[line, IndexColumnOfTime] = DateTime.MinValue.AddSeconds((line - IndexOfFirstLineVoltage) * _excelMeasurementsPeriod);
            _excelWorkSheet.Cells[IndexLineOfDuration, IndexColumnOfDuration] = _excelWorkSheet.Cells[line, IndexColumnOfTime];
        }

        private void TimerEventUartRxErrProcessor(object sender, ElapsedEventArgs eventArgs)
        {
            _uartRxErrTimer.Stop();
            ResetLastCommand();
        }

        private void LoopGetMeasurements(object obj)
        {
            if (_isPollEnable)
            {
                SendCommandToBms(UartCommand.UartCommandPcSendStatusRegisters);
            }
        }

        private void SendDataToBms(byte[] data, int size)
        {
            WriteTextToTextBoxTransmitMain(data, size);
            _serialPort.Write(data, 0, size);
        }

        public void SendCommandToBms(UartCommand command)
        {
            if (_lastCommand == UartCommand.Null)
            {
                _lastCommand = command;
                byte[] buffer = {(byte) command};
                WriteTextToTextBoxTransmitMain(buffer, buffer.Length);
                _uartRxErrTimer.Start();
                _serialPort.Write(buffer, 0, buffer.Length);
                _form.setPictureBoxRT_status(MemoryBms.RT_status.Transmit);
            } else
            {
                _unsentСommands.Enqueue(command);
            }
        }

        public void ResetLastCommand()
        {
            _countOfReceiveBuffer = 0;
            _lastCommand = UartCommand.Null;
            SendNextCommand();
        }

        private void SendNextCommand()
        {
            if (_unsentСommands.Count > 0)
            {
                SendCommandToBms(_unsentСommands.Dequeue());
            }
        }

        private string ConvertBytesToVerticalText(byte[] buffer, int size)
        {
            string str = null;
            for (int i = 0; i < size; i++)
            {
                str += String.Format($"{buffer[i]:X2}\r\n");
            }
            return str;
        }

        private void WriteTextToTextBoxTransmitMain(byte[] buffer, int size)
        {
            string str = ConvertBytesToVerticalText(buffer, size);
            _form.AddTextToTextBoxTransmitMain(str);
        }

        private void ComPortDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            const int inBufferSize = 128;
            byte[] inBuffer = new byte[inBufferSize];
            int rxCount = _serialPort.Read(inBuffer, 0, inBufferSize);
            string str = ConvertBytesToVerticalText(inBuffer, rxCount);
            _form.AddTextToTextBoxReceiveMain(str);
            ParseReseivedAnswer(inBuffer, rxCount);
        }

        private void ParseReseivedAnswer(byte[] inBuffer, int size)
        {
            switch (_lastCommand)
            {
                case UartCommand.UartCommandPcSendStatusRegisters:
                    for (int i = 0; (i < size) && (_countOfReceiveBuffer < _receiveBufferStatusRegisters.Length); i++)
                    {
                        _receiveBufferStatusRegisters[_countOfReceiveBuffer++] = inBuffer[i];
                        if ((_countOfReceiveBuffer == _receiveBufferStatusRegisters.Length) 
                            && (_receiveBufferStatusRegisters[_countOfReceiveBuffer - 1] == CalculateChecksum(_receiveBufferStatusRegisters, _receiveBufferStatusRegisters.Length - 1)))
                        {
                            _uartRxErrTimer.Stop();
                            _memory.StatusRegisters = StructTools.RawDeserialize<StatusRegisters>(_receiveBufferStatusRegisters, 0);
                            _form.UpdateMeasurements();
                            if (IsWriteToExcel)
                            {
                                UpdateMeasurementsToExcel();
                            }
                            _form.setPictureBoxRT_status(MemoryBms.RT_status.Receive);
                            ResetLastCommand();
                            break;
                        }
                    }
                    break;
                case UartCommand.UartCommandPcCalibrationU1:
                    if (Calibration(MemoryBms.CellIndex.Cell_1, inBuffer) == true)
                    {
                        ResetLastCommand();
                    }
                    break;
                case UartCommand.UartCommandPcCalibrationU2:
                    if (Calibration(MemoryBms.CellIndex.Cell_2, inBuffer) == true)
                    {
                        ResetLastCommand();
                    }
                    break;
                case UartCommand.UartCommandPcCalibrationU3:
                    if (Calibration(MemoryBms.CellIndex.Cell_3, inBuffer) == true)
                    {
                        ResetLastCommand();
                    }
                    break;
                case UartCommand.UartCommandPcCalibrationU4:
                    if (Calibration(MemoryBms.CellIndex.Cell_4, inBuffer) == true)
                    {
                        ResetLastCommand();
                    }
                    break;
                case UartCommand.UartCommandPcCalibrationTemp:
                    if (Calibration(MemoryBms.CellIndex.Temperature, inBuffer) == true)
                    {
                        ResetLastCommand();
                    }
                    break;
                case UartCommand.UartCommandPcSendSettingsCharge:
                    for (int i = 0; i < size; i++)
                    {
                        _bufferSettingsCharge[_countOfReceiveBuffer++] = inBuffer[i];
                        if (_countOfReceiveBuffer == _bufferSettingsCharge.Length)
                        {
                            _memory.SettingsCharge = StructTools.RawDeserialize<SettingsCharge>(_bufferSettingsCharge, 0);
                            _form.UpdateSettingsCharge();
                            ResetLastCommand();
                            break;
                        }
                    }
                    break;
                case UartCommand.UartCommandPcSendSettingsAlarm:
                    for (int i = 0; i < size; i++)
                    {
                        _bufferSettingsAlarm[_countOfReceiveBuffer++] = inBuffer[i];
                        if (_countOfReceiveBuffer == _bufferSettingsAlarm.Length)
                        {
                            _memory.SettingsAlarm = StructTools.RawDeserialize<SettingsAlarm>(_bufferSettingsAlarm, 0);
                            _form.UpdateSettingsAlarm();
                            ResetLastCommand();
                            break;
                        }
                    }
                    break;
                case UartCommand.UartCommandPcUpdateSettingsCharge:
                    if (UpdateSettings(inBuffer, _bufferSettingsCharge, StructTools.RawSerialize(_memory.SettingsCharge)) == true)
                    {
                        ResetLastCommand();
                    }
                    break;
                case UartCommand.UartCommandPcUpdateSettingsAlarm:
                    if (UpdateSettings(inBuffer, _bufferSettingsAlarm, StructTools.RawSerialize(_memory.SettingsAlarm)) == true)
                    {
                        ResetLastCommand();
                    }
                    break;
                default:
                    ResetLastCommand();
                    break;
            }
        }

        private bool Calibration(MemoryBms.CellIndex index, byte[] buffer)
        {
            bool finish = false;
            switch (_step) 
            {
                case 0:
                    if (buffer[0] == (byte)UartCommand.UartCommandOk)
                    {
                        byte[] txBuffer = (index == MemoryBms.CellIndex.Temperature) 
                            ? new[] { (byte) _calibrationValue} 
                            : new[] { (byte)(_calibrationValue >> 8), (byte)_calibrationValue };
                        SendDataToBms(txBuffer, txBuffer.Length);
                        _step++;
                    } else
                    {
                        _step = 0;
                        finish = true;
                    }
                    break;
                case 1:
                    ushort recaivedCalibrationValue = (index == MemoryBms.CellIndex.Temperature)
                        ? buffer[0]
                        : (ushort)((buffer[0] << 8) | buffer[1]);
                    if (recaivedCalibrationValue == _calibrationValue)
                    {
                        byte[] txBuffer = { (byte)UartCommand.UartCommandOk };
                        SendDataToBms(txBuffer, txBuffer.Length);
                        _step++;
                    } else
                    {
                        _step = 0;
                        finish = true;
                    }
                    break;
                case 2:
                    _calibrationCoefficient = (uint) ((buffer[0] << 24) | buffer[1] << 16 | buffer[2] << 8 | buffer[3]);
                    _form.UpdateTextInCalibrationK(index);
                    _step = 0;
                    finish = true;
                    break;
                default:
                    _step = 0;
                    finish = true;
                    break;
            }
            return finish;
        }

        //private bool CalibrationT(MemoryBms.CellIndex index, byte[] buffer, int size)
        //{
        //    bool finish = false;
        //    switch (_step)
        //    {
        //        case 0:
        //            if (buffer[0] == (byte)UartCommand.UartCommandOk)
        //            {
        //                byte[] txBuffer = { (byte)(_calibrationValue >> 8), (byte)_calibrationValue };
        //                SendDataToBms(txBuffer, txBuffer.Length);
        //                _step++;
        //            }
        //            else
        //            {
        //                _step = 0;
        //                finish = true;
        //            }
        //            break;
        //        case 1:
        //            ushort calibrationU = (ushort)((buffer[0] << 8) | buffer[1]);
        //            if (calibrationU == _calibrationValue)
        //            {
        //                byte[] txBuffer = { (byte)UartCommand.UartCommandOk };
        //                SendDataToBms(txBuffer, txBuffer.Length);
        //                _step++;
        //            }
        //            else
        //            {
        //                _step = 0;
        //                finish = true;
        //            }
        //            break;
        //        case 2:
        //            _calibrationCoefficient = (uint)((buffer[0] << 24) | buffer[1] << 16 | buffer[2] << 8 | buffer[3]);
        //            _form.UpdateTextInCalibrationK(index);
        //            _step = 0;
        //            finish = true;
        //            break;
        //        default:
        //            _step = 0;
        //            finish = true;
        //            break;
        //    }
        //    return finish;
        //}

        public byte CalculateChecksum(byte[] array, int length)
        {
            ushort checksum = 0;
            for (int i = 0; i < length; i++)
            {
                checksum += (ushort)(array[i] * ChecksumConstant * (i + 1));
            }
            return (byte)checksum;
        }

        private bool UpdateSettings(byte[] buffer, byte[] bufferSettings, byte[] settingsStruct)
        {
            bool finish = false;
            switch (_step) 
            {
                case 0:
                    if (buffer[0] == (byte)UartCommand.UartCommandOk)
                    {
                        Array.Copy(settingsStruct, bufferSettings, bufferSettings.Length);
                        SendDataToBms(bufferSettings, bufferSettings.Length);
                        _step++;
                    }
                    else
                    {
                        _step = 0;
                        finish = true;
                    }
                    break;
                case 1:
                    if (buffer[0] == CalculateChecksum(bufferSettings, bufferSettings.Length))
                    {
                        byte[] txBuffer = { (byte)UartCommand.UartCommandOk };
                        SendDataToBms(txBuffer, txBuffer.Length);
                    }
                    _step = 0;
                    finish = true;
                    break;
                default:
                    _step = 0;
                    finish = true;
                    break;
            }
            return finish;
        }

        private void MainLoop(object obj)
        {
            if (_isPollEnable)
            {
                
            }
        }

        public void SetMeasurementsPeriod(int period)
        {
            if (_isPollEnable == false)
            {
                _pollingPeriod = period;
                _mainLoopPeriod = _pollingPeriod / 10;
                UpdateTimers();
            }
        }

        public void SetExcelMeasurementsPeriod(int period)
        {
            if (_isPollEnable == false)
            {
                _excelMeasurementsPeriod = period;
            }
        }

        private void UpdateTimers()
        {
            _timerLoopGetMeasurements.Change(0, _pollingPeriod);
            _timerMainLoop.Change(0, _mainLoopPeriod);
        }

        public bool IsPollEnable
        {
            set => _isPollEnable = value;
        }

        public bool IsWriteToExcel { private get; set; } = false;

        public MemoryBms GetMemory()
        {
            return _memory;
        }

        public void SetCalibrationValue(ushort value)
        {
            _calibrationValue = value;
        }

        public uint GetCalibrationCoefficient()
        {
            return _calibrationCoefficient;
        }

        public void FreeingUpResources()
        {
            _uartRxErrTimer.Dispose();
        }
    }
}
