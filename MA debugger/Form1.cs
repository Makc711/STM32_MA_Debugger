using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Threading;
using System.Windows.Forms;

namespace MA_debugger
{
    public partial class Form1 : Form
    {
        private BatteryModule _batteryModule;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SearchSerialPorts();
            _batteryModule = new BatteryModule(this, serialPort1, Int32.Parse(textBoxMeasurementPeriod.Text));
        }

        private void buttonConnect_Click(object sender, EventArgs e)
        {
            if (buttonConnect.Text == @"Connect")
            {
                try
                {
                    ConfigureSerialPort();
                    serialPort1.Open();
                    buttonConnect.Text = @"Disconnect";
                    buttonReScan.Enabled = false;
                } catch(Exception ex)
                {
                    MessageBox.Show(ex.Message, @"Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                buttonConnect.Text = @"Connect";
                buttonReScan.Enabled = true;
                CloseSerialPort();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            CloseSerialPort();
        }

        private void CloseSerialPort()
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.Close();
            }
        }

        private void SearchSerialPorts()
        {
            comboBoxCOMs.Items.Clear();
            string[] ports = SerialPort.GetPortNames();
            List<string> sortingPorts = new List<string>();
            sortingPorts.AddRange(ports);
            sortingPorts.Sort();
            foreach (var port in sortingPorts)
            {
                if (port.StartsWith("COM"))
                {
                    comboBoxCOMs.Items.Add(port);
                }
            }
            if (comboBoxCOMs.Items.Count > 0)
            {
                comboBoxCOMs.SelectedIndex = 0;
            } else
            {
                comboBoxCOMs.Text = string.Empty;
            }
        }

        private void buttonReScan_Click(object sender, EventArgs e)
        {
            SearchSerialPorts();
        }

        private void ConfigureSerialPort()
        {
            serialPort1.PortName = comboBoxCOMs.Text;
            serialPort1.BaudRate = GetSelectedBaudRate();
            serialPort1.DataBits = GetSelectedDataBits();
            serialPort1.Parity = GetSelectedParity();
            serialPort1.StopBits = GetSelectedStopBits();
            serialPort1.Handshake = GetSelectedHandshake();
        }

        private int GetSelectedBaudRate()
        {
            int baudRate = 9600;
            foreach (RadioButton radioButton in groupBoxBR.Controls)
            {
                if (radioButton.Checked)
                {
                    baudRate = Int32.Parse(radioButton.Text);
                    break;
                }
            }
            return baudRate;
        }

        private int GetSelectedDataBits()
        {
            int dataBits = 8;
            foreach (RadioButton radioButton in groupBoxDB.Controls)
            {
                if (radioButton.Checked)
                {
                    dataBits = Int32.Parse(radioButton.Text);
                    break;
                }
            }
            return dataBits;
        }

        private Parity GetSelectedParity()
        {
            Parity parity = Parity.None;
            if (radioButtonParityNone.Checked)
            {
                parity = Parity.None;
            } else if (radioButtonParityOdd.Checked)
            {
                parity = Parity.Odd;
            } else if (radioButtonParityEven.Checked)
            {
                parity = Parity.Even;
            } else if (radioButtonParityMark.Checked)
            {
                parity = Parity.Mark;
            } else if (radioButtonParitySpace.Checked)
            {
                parity = Parity.Space;
            }
            return parity;
        }

        private StopBits GetSelectedStopBits()
        {
            StopBits stopBits = StopBits.One;
            if (radioButtonSB1.Checked)
            {
                stopBits = StopBits.One;
            } else if (radioButtonSB15.Checked)
            {
                stopBits = StopBits.OnePointFive;
            } else if (radioButtonSB2.Checked)
            {
                stopBits = StopBits.Two;
            }
            return stopBits;
        }

        private Handshake GetSelectedHandshake()
        {
            Handshake handshake = Handshake.None;
            if (radioButtonHSNone.Checked)
            {
                handshake = Handshake.None;
            } else if (radioButtonHSRC.Checked)
            {
                handshake = Handshake.RequestToSend;
            } else if (radioButtonHSOO.Checked)
            {
                handshake = Handshake.XOnXOff;
            } else if (radioButtonHSRCOO.Checked)
            {
                handshake = Handshake.RequestToSendXOnXOff;
            }
            return handshake;
        }

        private void buttonMeasurementsEnable_Click(object sender, EventArgs e)
        {
            buttonMeasurementsDisable.Enabled = true;
            buttonMeasurementsEnable.Enabled = false;
            _batteryModule.SetMeasurementsPeriod(Int32.Parse(textBoxMeasurementPeriod.Text));
            _batteryModule.IsMeasurementsEnable = true;
        }

        private void buttonMeasurementsDisable_Click(object sender, EventArgs e)
        {
            buttonMeasurementsEnable.Enabled = true;
            buttonMeasurementsDisable.Enabled = false;
            _batteryModule.IsMeasurementsEnable = false;
        }

        public void AddTextToTextBoxTransmit(string text)
        {
            textBoxTransmit.Invoke((ThreadStart) delegate
            {
                textBoxTransmit.AppendText(text);
            });
        }

        public void AddTextToTextBoxReceive(string text)
        {
            textBoxReceive.Invoke((ThreadStart) delegate
            {
                textBoxReceive.AppendText(text);
            });
        }

        public void UpdateMeasurements()
        {
            textBoxUcell.Invoke((ThreadStart)delegate
            {
                textBoxUcell.Text = _batteryModule.GetMemory().Measurements.U_cell.ToString();
            });
            textBoxIbalance.Invoke((ThreadStart)delegate
            {
                textBoxIbalance.Text = _batteryModule.GetMemory().Measurements.I_balance.ToString();
            });
            textBoxTemperatureAnode.Invoke((ThreadStart)delegate
            {
                textBoxTemperatureAnode.Text = _batteryModule.GetMemory().Measurements.TemperatureAnode.ToString();
            });
            textBoxTemperatureCathode.Invoke((ThreadStart)delegate
            {
                textBoxTemperatureCathode.Text = _batteryModule.GetMemory().Measurements.TemperatureCathode.ToString();
            });
            textBoxTemperatureVT1.Invoke((ThreadStart)delegate
            {
                textBoxTemperatureVT1.Text = _batteryModule.GetMemory().Measurements.TemperatureVT1.ToString();
            });

            UpdateEventRegister();
        }

        private void UpdateEventRegister()
        {
            ushort maEventRegister = _batteryModule.GetMemory().Measurements.MA_Event_Register;
            pictureBuffPwrNormal.Invoke((ThreadStart)delegate
            {
                pictureBuffPwrNormal.BackColor = (maEventRegister & (ushort)MemoryMa.MaEvent.BufferEnableMsk) == (ushort)MemoryMa.MaEvent.BufferEnable
                    ? pictureBoxGreen.BackColor
                    : pictureBoxRed.BackColor;
            });
            pictureTransformerOUT.Invoke((ThreadStart)delegate
            {
                pictureTransformerOUT.BackColor = (maEventRegister & (ushort)MemoryMa.MaEvent.TransformerOutMsk) == (ushort)MemoryMa.MaEvent.TransformerOut
                    ? pictureBoxGreen.BackColor
                    : pictureBoxRed.BackColor;
            });
            pictureBalancingInEnable.Invoke((ThreadStart)delegate
            {
                pictureBalancingInEnable.BackColor = (maEventRegister & (ushort)MemoryMa.MaEvent.BalancingInMsk) == (ushort)MemoryMa.MaEvent.BalancingIn
                    ? pictureBoxGreen.BackColor
                    : pictureBoxRed.BackColor;
            });
            pictureBalancingOutEnable.Invoke((ThreadStart)delegate
            {
                pictureBalancingOutEnable.BackColor = (maEventRegister & (ushort)MemoryMa.MaEvent.BalancingOutMsk) == (ushort)MemoryMa.MaEvent.BalancingOut
                    ? pictureBoxGreen.BackColor
                    : pictureBoxRed.BackColor;
            });
            pictureCellOvervoltage.Invoke((ThreadStart)delegate
            {
                pictureCellOvervoltage.BackColor = (maEventRegister & (ushort)MemoryMa.MaEvent.SafetyStatusCovMsk) == (ushort)MemoryMa.MaEvent.SafetyStatusCov
                    ? pictureBoxGreen.BackColor
                    : pictureBoxRed.BackColor;
            });
            pictureCellUndervoltage.Invoke((ThreadStart)delegate
            {
                pictureCellUndervoltage.BackColor = (maEventRegister & (ushort)MemoryMa.MaEvent.SafetyStatusCuvMsk) == (ushort)MemoryMa.MaEvent.SafetyStatusCuv
                    ? pictureBoxGreen.BackColor
                    : pictureBoxRed.BackColor;
            });
            pictureCellOvertemperature.Invoke((ThreadStart)delegate
            {
                pictureCellOvertemperature.BackColor = (maEventRegister & (ushort)MemoryMa.MaEvent.SafetyStatusCotMsk) == (ushort)MemoryMa.MaEvent.SafetyStatusCot
                    ? pictureBoxGreen.BackColor
                    : pictureBoxRed.BackColor;
            });
            pictureCellUndertemperature.Invoke((ThreadStart)delegate
            {
                pictureCellUndertemperature.BackColor = (maEventRegister & (ushort)MemoryMa.MaEvent.SafetyStatusCutMsk) == (ushort)MemoryMa.MaEvent.SafetyStatusCut
                    ? pictureBoxGreen.BackColor
                    : pictureBoxRed.BackColor;
            });
            pictureTransistorOvertemperature.Invoke((ThreadStart)delegate
            {
                pictureTransistorOvertemperature.BackColor = (maEventRegister & (ushort)MemoryMa.MaEvent.SafetyStatusOttMsk) == (ushort)MemoryMa.MaEvent.SafetyStatusOtt
                    ? pictureBoxGreen.BackColor
                    : pictureBoxRed.BackColor;
            });
            pictureMaCircuitError.Invoke((ThreadStart)delegate
            {
                pictureMaCircuitError.BackColor = (maEventRegister & (ushort)MemoryMa.MaEvent.SafetyStatusMaFailMsk) == (ushort)MemoryMa.MaEvent.SafetyStatusMaFail
                    ? pictureBoxGreen.BackColor
                    : pictureBoxRed.BackColor;
            });
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            textBoxReceive.Text = string.Empty;
            textBoxTransmit.Text = string.Empty;
            ClearMeasurementsFields();
        }

        private void ClearMeasurementsFields()
        {
            textBoxUcell.Text = string.Empty;
            textBoxIbalance.Text = string.Empty;
            textBoxTemperatureAnode.Text = string.Empty;
            textBoxTemperatureCathode.Text = string.Empty;
            textBoxTemperatureVT1.Text = string.Empty;
            pictureBuffPwrNormal.BackColor = Color.Gainsboro;
            pictureTransformerOUT.BackColor = Color.Gainsboro;
            pictureBalancingInEnable.BackColor = Color.Gainsboro;
            pictureBalancingOutEnable.BackColor = Color.Gainsboro;
            pictureCellOvervoltage.BackColor = Color.Gainsboro;
            pictureCellUndervoltage.BackColor = Color.Gainsboro;
            pictureCellOvertemperature.BackColor = Color.Gainsboro;
            pictureCellUndertemperature.BackColor = Color.Gainsboro;
            pictureTransistorOvertemperature.BackColor = Color.Gainsboro;
            pictureMaCircuitError.BackColor = Color.Gainsboro;
            pictureSendSettings.BackColor = Color.Gainsboro;
            pictureCheckSettings.BackColor = Color.Gainsboro;
        }

        private void buttonBalanceOUT_Click(object sender, EventArgs e)
        {
            _batteryModule.SendCommandToMa(BatteryModule.UartCommand.MkBalancingOutEnable);
        }

        private void buttonBalanceIN_Click(object sender, EventArgs e)
        {
            _batteryModule.SendCommandToMa(BatteryModule.UartCommand.MkBalancingInEnable);
        }

        private void buttonStopBalance_Click(object sender, EventArgs e)
        {
            _batteryModule.SendCommandToMa(BatteryModule.UartCommand.MkBalancingDisable);
        }

        private void buttonSendSettings_Click(object sender, EventArgs e)
        {
            SaveSettingsToMemory();
            _batteryModule.SetStatusWaitSettings();
        }

        private void SaveSettingsToMemory()
        {
            MaSettings settings;
            settings.COV_Threshold = (ushort) int.Parse(textBoxCOV_Threshold.Text);
            settings.COV_Recovery = (ushort)int.Parse(textBoxCOV_Recovery.Text);
            settings.COV_Time = (ushort)int.Parse(textBoxCOV_Time.Text);
            settings.CUV_Threshold = (ushort)int.Parse(textBoxCUV_Threshold.Text);
            settings.CUV_Recovery = (ushort)int.Parse(textBoxCUV_Recovery.Text);
            settings.CUV_Time = (ushort)int.Parse(textBoxCUV_Time.Text);
            settings.COT_Threshold = (sbyte) int.Parse(textBoxCOT_Threshold.Text);
            settings.COT_Recovery = (sbyte)int.Parse(textBoxCOT_Recovery.Text);
            settings.COT_Time = (ushort)int.Parse(textBoxCOT_Time.Text);
            settings.CUT_Threshold = (sbyte)int.Parse(textBoxCUT_Threshold.Text);
            settings.CUT_Recovery = (sbyte)int.Parse(textBoxCUT_Recovery.Text);
            settings.CUT_Time = (ushort)int.Parse(textBoxCUT_Time.Text);
            settings.OTT_Threshold = (sbyte)int.Parse(textBoxOTT_Threshold.Text);
            settings.OTT_Recovery = (sbyte)int.Parse(textBoxOTT_Recovery.Text);
            settings.OTT_Time = (ushort)int.Parse(textBoxOTT_Time.Text);
            _batteryModule.GetMemory().Settings = settings;
        }

        public void SetSendSettingsStatus(bool isSentWithoutErrors)
        {
            pictureSendSettings.Invoke((ThreadStart)delegate
            {
                pictureSendSettings.BackColor = isSentWithoutErrors ? pictureBoxGreen.BackColor : pictureBoxRed.BackColor;
            });
        }

        public void SetChecksumStatus(bool isSentWithoutErrors)
        {
            pictureCheckSettings.Invoke((ThreadStart)delegate
            {
                pictureCheckSettings.BackColor = isSentWithoutErrors ? pictureBoxGreen.BackColor : pictureBoxRed.BackColor;
            });
        }

        private void buttonCheckSettings_Click(object sender, EventArgs e)
        {
            SaveSettingsToMemory();
            _batteryModule.GetMemory().SettingsBuffer = StructTools.RawSerialize(_batteryModule.GetMemory().Settings);
            _batteryModule.SendCommandToMa(BatteryModule.UartCommand.MkSendSettingsChecksum);
        }
    }
}
