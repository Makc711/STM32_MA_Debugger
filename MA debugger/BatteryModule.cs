using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;

namespace MA_debugger
{
    class BatteryModule
    {
        private int _mainLoopPeriod;
        private int _measurementsPeriod;
        private readonly MemoryMa _memory = new MemoryMa();
        private readonly Form1 _form;
        private readonly SerialPort _serialPort;
        private readonly Timer _timerMainLoop;
        private readonly Timer _timerLoopGetMeasurements;
        private bool _isMeasurementsEnable; // = false
        private bool _isMeasurementsWasEnabled; // = false
        private UartCommand _lastCommand = UartCommand.Null;
        private readonly byte[] _receiveBuffer = new byte[Marshal.SizeOf(typeof(MaMeasurements))];
        private int _countOfReceiveBufferBytes; // = 0
        private readonly Queue<UartCommand> _unsentСommands = new Queue<UartCommand>();
        private int _lastSentSettingsAddress; // = 0
        private bool _isSettingsStartedSend; // = false
        private const int PackageSize = 8;
        private const int CommandSize = 1;
        private const int MaxBytesInPackage = PackageSize - CommandSize;
        private const int SizeOfSendPackageSize = 1;

        internal enum UartCommand : byte
        {
            Null                     = 0x00,
            MkBalancingInEnable      = 0x51, /*!< Command from MK. Enable balancing in. */
            MkBalancingOutEnable     = 0x52, /*!< Command from MK. Enable balancing out. */
            MkBalancingDisable       = 0x53, /*!< Command from MK. Disable balancing. */
            MkSendMeasurements       = 0x54, /*!< Command from MK. Send the measurements struct. */
            MkWaitSettings           = 0x55, /*!< Command from MK. Wait settings. */
            MkSendSettingsChecksum   = 0x56, /*!< Command from MK. Send checksum settings. */
            MkUpdateSettings         = 0x57, /*!< Command from MK. Update settings. */
            MkOk                     = 0x5F, /*!< Command from MK. OK. */
            MkError                  = 0x50, /*!< Command from MK. Error. */
            MaOk                     = 0x2F, /*!< Command from MA. OK. */
            MaError                  = 0x20  /*!< Command from MA. Error. */
        }

        private enum UartCommandState : byte
        {
            DataPos = 7,                              
            DataMsk = 1 << DataPos   /*!< 0x80, Command_State Data*/
        }

        public BatteryModule(Form1 form, SerialPort serialPort, int measurementsPeriod)
        {
            _form = form;
            _serialPort = serialPort;
            _measurementsPeriod = measurementsPeriod;
            _mainLoopPeriod = _measurementsPeriod / 10;
            _timerLoopGetMeasurements = new Timer(LoopGetMeasurements, null, 0, _measurementsPeriod);
            _timerMainLoop = new Timer(MainLoop, null, 0, _mainLoopPeriod);
            _serialPort.DataReceived += ComPortDataReceived;
        }

        private void LoopGetMeasurements(object obj)
        {
            if (_isMeasurementsEnable)
            {
                SendCommandToMa(UartCommand.MkSendMeasurements);
            }
        }

        public void SendCommandToMa(UartCommand command)
        {
            if (_lastCommand == UartCommand.Null)
            {
                _lastCommand = command;
                byte[] buffer = {1, (byte) command};
                AddTextToTransmitTextBox(buffer, buffer.Length);
                _serialPort.Write(buffer, 0, buffer.Length);
            } else
            {
                _unsentСommands.Enqueue(command);
            }
        }

        private void SendNextCommand()
        {
            if (_unsentСommands.Count > 0)
            {
                SendCommandToMa(_unsentСommands.Dequeue());
            }
        }

        public void SetStatusWaitSettings()
        {
            _isMeasurementsWasEnabled = _isMeasurementsEnable;
            _isMeasurementsEnable = false;
            _memory.SettingsBuffer = StructTools.RawSerialize(_memory.Settings);
            SendCommandToMa(UartCommand.MkWaitSettings);
        }

        private void SendSettings()
        {
            int countOfSentBytes;
            if (_memory.SettingsBuffer.Length - _lastSentSettingsAddress < MaxBytesInPackage)
            {
                countOfSentBytes = _memory.SettingsBuffer.Length - _lastSentSettingsAddress;
            } else
            {
                countOfSentBytes = MaxBytesInPackage;
            }
            SendData(_memory.SettingsBuffer, _lastSentSettingsAddress, countOfSentBytes);
        }

        private void SendData(byte[] data, int startDataAddress, int size)
        {
            int sendPackageSize = size + CommandSize;
            if (sendPackageSize > PackageSize)
                throw new ArgumentException("Message size is larger than allowed. Message size = " + sendPackageSize + ", Package size = " + PackageSize);
            byte address = (byte) ((byte) UartCommandState.DataMsk | startDataAddress);
            byte[] buffer = new byte[sendPackageSize + SizeOfSendPackageSize];
            buffer[0] = (byte) sendPackageSize;
            buffer[1] = address;
            Array.Copy(data, startDataAddress, buffer, 2, size);
            AddTextToTransmitTextBox(buffer, buffer.Length);
            _serialPort.Write(buffer, 0, buffer.Length);
        }

        private void AddTextToTransmitTextBox(byte[] buffer, int size)
        {
            string str = null;
            for (int i = 0; i < size; i++)
            {
                str += String.Format($"{buffer[i]:X2}\r\n");
            }
            _form.AddTextToTextBoxTransmit(str);
        }

        private void ComPortDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            const int inBufferSize = 128;
            byte[] inBuffer = new byte[inBufferSize];
            int rxCount = _serialPort.Read(inBuffer, 0, inBufferSize);
            string str = null;
            for (int i = 0; i < rxCount; i++)
            {
                str += String.Format($"{inBuffer[i]:X2}\r\n");
            }
            _form.AddTextToTextBoxReceive(str);
            ParseReseivedAnswer(inBuffer, rxCount);
        }

        private void ParseReseivedAnswer(byte[] inBuffer, int size)
        {
            switch (_lastCommand)
            {
                case UartCommand.MkSendMeasurements:
                    for (int i = 0; i < size; i++)
                    {
                        if (_countOfReceiveBufferBytes == _receiveBuffer.Length)
                        {
                            if ((inBuffer[i] == (byte) UartCommand.MaOk) && (i == size - 1))
                            {
                                _memory.Measurements = StructTools.RawDeserialize<MaMeasurements>(_receiveBuffer, 0);
                                _form.UpdateMeasurements();
                            } else
                            {
                                // Если 3 раза ответ неверный, то записать в журнал "Ошибка связи"
                            }
                            _countOfReceiveBufferBytes = 0;
                            _lastCommand = UartCommand.Null;
                            SendNextCommand();
                            break;
                        }
                        _receiveBuffer[_countOfReceiveBufferBytes++] = inBuffer[i];
                    }
                    break;
                case UartCommand.MkBalancingOutEnable:
                    _lastCommand = UartCommand.Null;
                    if ((inBuffer[0] == (byte) UartCommand.MaOk) && (size == 1))
                    {
                        SendNextCommand();
                    } else
                    {
                        SendCommandToMa(UartCommand.MkBalancingOutEnable); // Только 3 попытки!
                    }
                    break;
                case UartCommand.MkBalancingInEnable:
                    _lastCommand = UartCommand.Null;
                    if ((inBuffer[0] == (byte) UartCommand.MaOk) && (size == 1))
                    {
                        SendNextCommand();
                    } else
                    {
                        SendCommandToMa(UartCommand.MkBalancingInEnable); // Только 3 попытки!
                    }
                    break;
                case UartCommand.MkBalancingDisable:
                    _lastCommand = UartCommand.Null;
                    if ((inBuffer[0] == (byte) UartCommand.MaOk) && (size == 1))
                    {
                        SendNextCommand();
                    } else
                    {
                        SendCommandToMa(UartCommand.MkBalancingDisable); // Только 3 попытки!
                    }
                    break;
                case UartCommand.MkWaitSettings:
                    if ((inBuffer[0] == (byte) UartCommand.MaOk) && (size == 1))
                    {
                        if (_isSettingsStartedSend == false)
                        {
                            _lastSentSettingsAddress = 0;
                            _isSettingsStartedSend = true;
                        } else
                        {
                            _lastSentSettingsAddress += MaxBytesInPackage;
                        }
                        if (_lastSentSettingsAddress < _memory.SettingsBuffer.Length)
                        {
                            SendSettings();
                        } else
                        {
                            _lastSentSettingsAddress = 0;
                            _lastCommand = UartCommand.Null;
                            SendCommandToMa(UartCommand.MkSendSettingsChecksum);
                        }
                    }
                    else
                    {
                        if (_isSettingsStartedSend)
                        {
                            SendSettings(); // Только 3 попытки!
                        } else
                        {
                            _lastCommand = UartCommand.Null;
                            SendCommandToMa(UartCommand.MkWaitSettings); // Только 3 попытки!
                        }
                    }
                    break;
                case UartCommand.MkSendSettingsChecksum:
                    _lastCommand = UartCommand.Null;
                    byte settingsChecksum = _memory.CalculateSettingsChecksum();
                    if ((inBuffer[0] == settingsChecksum) && (inBuffer[1] == (byte) UartCommand.MaOk) && (size == 2))
                    {
                        _form.SetChecksumStatus(true);
                        if (_isSettingsStartedSend)
                        {
                            SendCommandToMa(UartCommand.MkUpdateSettings);
                        } else
                        {
                            SendNextCommand();
                        }
                    } else
                    {
                        _form.SetChecksumStatus(false);
                        SendNextCommand();
                        EndSendSettings();
                    }
                    break;
                case UartCommand.MkUpdateSettings:
                    _lastCommand = UartCommand.Null;
                    _form.SetSendSettingsStatus((inBuffer[0] == (byte) UartCommand.MaOk) && (size == 1));
                    SendNextCommand();
                    EndSendSettings();
                    break;
                default:
                    _lastCommand = UartCommand.Null;
                    break;
            }
        }

        private void EndSendSettings()
        {
            if (_isSettingsStartedSend)
            {
                _isSettingsStartedSend = false;
                _isMeasurementsEnable = _isMeasurementsWasEnabled;
            }
        }

        private void MainLoop(object obj)
        {
            if (_isMeasurementsEnable)
            {
                
            }
        }

        public void SetMeasurementsPeriod(int period)
        {
            if (_isMeasurementsEnable == false)
            {
                _measurementsPeriod = period;
                _mainLoopPeriod = _measurementsPeriod / 10;
                UpdateTimers();
            }
        }

        private void UpdateTimers()
        {
            _timerLoopGetMeasurements.Change(0, _measurementsPeriod);
            _timerMainLoop.Change(0, _mainLoopPeriod);
        }

        public bool IsMeasurementsEnable
        {
            set => _isMeasurementsEnable = value;
        }

        public MemoryMa GetMemory()
        {
            return _memory;
        }
    }
}
