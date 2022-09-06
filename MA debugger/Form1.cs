using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO.Ports;
using System.Security.Permissions;
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
            ClearMeasurementsFields();
            _batteryModule = new BatteryModule(this, serialPort1, Int32.Parse(textBoxPollingPeriod.Text));
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
            _batteryModule.FreeingUpResources();
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

        private void buttonPollEnable_Click(object sender, EventArgs e)
        {
            buttonPollDisable.Enabled = true;
            buttonPollEnable.Enabled = false;
            checkBoxWriteToExcel.Enabled = false;
            if (checkBoxWriteToExcel.Checked)
            {
                _batteryModule.ExcelOpenBook();
                _batteryModule.IsAddStartDataToExcel = false;
            }
            _batteryModule.SetMeasurementsPeriod(Int32.Parse(textBoxPollingPeriod.Text));
            _batteryModule.SetExcelMeasurementsPeriod(Int32.Parse(textBoxExcelMeasurementsPeriod.Text));
            _batteryModule.ResetMeasurements();
            _batteryModule.IsPollEnable = true;
        }

        private void buttonPollDisable_Click(object sender, EventArgs e)
        {
            buttonPollEnable.Enabled = true;
            buttonPollDisable.Enabled = false;
            checkBoxWriteToExcel.Enabled = true;
            _batteryModule.IsPollEnable = false;
            if (checkBoxWriteToExcel.Checked)
            {
                _batteryModule.ExcelSaveBook();
            }
        }

        public void AddTextToTextBoxTransmitMain(string text)
        {
            textBoxTransmitMain.Invoke((ThreadStart) delegate
            {
                textBoxTransmitMain.AppendText(text);
            });
        }

        public void AddTextToTextBoxReceiveMain(string text)
        {
            textBoxReceiveMain.Invoke((ThreadStart) delegate
            {
                textBoxReceiveMain.AppendText(text);
            });
        }

        public void UpdateMeasurements()
        {
            textBoxBatteryVoltage.Invoke((ThreadStart)delegate
            {
                textBoxBatteryVoltage.Text = (decimal.ToDouble(_batteryModule.GetMemory().StatusRegisters.batteryVoltage) / 1000).ToString(CultureInfo.InvariantCulture);
            });
            textBoxCellVoltageMax.Invoke((ThreadStart)delegate
            {
                textBoxCellVoltageMax.Text = _batteryModule.GetMemory().StatusRegisters.cellVoltageMax.ToString();
            });
            textBoxCellVoltageMin.Invoke((ThreadStart)delegate
            {
                textBoxCellVoltageMin.Text = _batteryModule.GetMemory().StatusRegisters.cellVoltageMin.ToString();
            });
            textBoxRadiatorTemperature.Invoke((ThreadStart)delegate
            {
                textBoxRadiatorTemperature.Text = _batteryModule.GetMemory().StatusRegisters.radiatorTemperature.ToString();
            });
            textBoxSOC.Invoke((ThreadStart)delegate
            {
                textBoxSOC.Text = _batteryModule.GetMemory().StatusRegisters.stateOfCharge.ToString();
            });
            textBoxСellVoltage1.Invoke((ThreadStart)delegate
            {
                textBoxСellVoltage1.Text = _batteryModule.GetMemory().StatusRegisters.cellVoltage1.ToString();
            });
            textBoxСellVoltage2.Invoke((ThreadStart)delegate
            {
                textBoxСellVoltage2.Text = _batteryModule.GetMemory().StatusRegisters.cellVoltage2.ToString();
            });
            textBoxСellVoltage3.Invoke((ThreadStart)delegate
            {
                textBoxСellVoltage3.Text = _batteryModule.GetMemory().StatusRegisters.cellVoltage3.ToString();
            });
            textBoxСellVoltage4.Invoke((ThreadStart)delegate
            {
                textBoxСellVoltage4.Text = _batteryModule.GetMemory().StatusRegisters.cellVoltage4.ToString();
            });
            textBoxADCVoltage1.Invoke((ThreadStart)delegate
            {
                textBoxADCVoltage1.Text = _batteryModule.GetMemory().StatusRegisters.ADCVoltage1.ToString();
            });
            textBoxADCVoltage2.Invoke((ThreadStart)delegate
            {
                textBoxADCVoltage2.Text = _batteryModule.GetMemory().StatusRegisters.ADCVoltage2.ToString();
            });
            textBoxADCVoltage3.Invoke((ThreadStart)delegate
            {
                textBoxADCVoltage3.Text = _batteryModule.GetMemory().StatusRegisters.ADCVoltage3.ToString();
            });
            textBoxADCVoltage4.Invoke((ThreadStart)delegate
            {
                textBoxADCVoltage4.Text = _batteryModule.GetMemory().StatusRegisters.ADCVoltage4.ToString();
            });

            progressBarСellVoltage1.Invoke((ThreadStart)delegate
            {
                ushort cellVoltage = _batteryModule.GetMemory().StatusRegisters.cellVoltage1;
                ushort voltage = calculateCellVoltage(cellVoltage);
                progressBarСellVoltage1.Value = voltage;
            });
            progressBarСellVoltage2.Invoke((ThreadStart)delegate
            {
                ushort cellVoltage = _batteryModule.GetMemory().StatusRegisters.cellVoltage2;
                ushort voltage = calculateCellVoltage(cellVoltage);
                progressBarСellVoltage2.Value = voltage;
            });
            progressBarСellVoltage3.Invoke((ThreadStart)delegate
            {
                ushort cellVoltage = _batteryModule.GetMemory().StatusRegisters.cellVoltage3;
                ushort voltage = calculateCellVoltage(cellVoltage);
                progressBarСellVoltage3.Value = voltage;
            });
            progressBarСellVoltage4.Invoke((ThreadStart)delegate
            {
                ushort cellVoltage = _batteryModule.GetMemory().StatusRegisters.cellVoltage4;
                ushort voltage = calculateCellVoltage(cellVoltage);
                progressBarСellVoltage4.Value = voltage;
            });

            UpdateCellStatus();
            UpdateCellSOC();
            UpdateCellChargeAlert();
            UpdateCellChargeStatus();
            UpdateCellSafetyAlert();
            UpdateCellSafetyStatus();
            UpdateBatteryStatus();
        }

        ushort calculateCellVoltage(ushort voltage)
        {
            ushort value = voltage;
            if (voltage <= 2800)
            {
                value = 2800;
            }
            if (voltage >= 3700)
            {
                value = 3700;
            }
            return value;
        }

        private void UpdateCellStatus()
        {
            ushort cellStatus1 = _batteryModule.GetMemory().StatusRegisters.cellStatus1;
            ushort cellStatus2 = _batteryModule.GetMemory().StatusRegisters.cellStatus2;
            ushort cellStatus3 = _batteryModule.GetMemory().StatusRegisters.cellStatus3;
            ushort cellStatus4 = _batteryModule.GetMemory().StatusRegisters.cellStatus4;
            pictureBoxCOVC1.Invoke((ThreadStart)delegate
            {
                pictureBoxCOVC1.BackColor = (cellStatus1 & (ushort)MemoryBms.CellStatus.Cell_Status_COVC_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_COVC
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCOVC2.Invoke((ThreadStart)delegate
            {
                pictureBoxCOVC2.BackColor = (cellStatus2 & (ushort)MemoryBms.CellStatus.Cell_Status_COVC_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_COVC
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCOVC3.Invoke((ThreadStart)delegate
            {
                pictureBoxCOVC3.BackColor = (cellStatus3 & (ushort)MemoryBms.CellStatus.Cell_Status_COVC_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_COVC
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCOVC4.Invoke((ThreadStart)delegate
            {
                pictureBoxCOVC4.BackColor = (cellStatus4 & (ushort)MemoryBms.CellStatus.Cell_Status_COVC_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_COVC
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCOVT1.Invoke((ThreadStart)delegate
            {
                pictureBoxCOVT1.BackColor = (cellStatus1 & (ushort)MemoryBms.CellStatus.Cell_Status_COVT_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_COVT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCOVT2.Invoke((ThreadStart)delegate
            {
                pictureBoxCOVT2.BackColor = (cellStatus2 & (ushort)MemoryBms.CellStatus.Cell_Status_COVT_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_COVT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCOVT3.Invoke((ThreadStart)delegate
            {
                pictureBoxCOVT3.BackColor = (cellStatus3 & (ushort)MemoryBms.CellStatus.Cell_Status_COVT_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_COVT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCOVT4.Invoke((ThreadStart)delegate
            {
                pictureBoxCOVT4.BackColor = (cellStatus4 & (ushort)MemoryBms.CellStatus.Cell_Status_COVT_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_COVT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCUV1.Invoke((ThreadStart)delegate
            {
                pictureBoxCUV1.BackColor = (cellStatus1 & (ushort)MemoryBms.CellStatus.Cell_Status_CUV_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CUV
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCUV2.Invoke((ThreadStart)delegate
            {
                pictureBoxCUV2.BackColor = (cellStatus2 & (ushort)MemoryBms.CellStatus.Cell_Status_CUV_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CUV
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCUV3.Invoke((ThreadStart)delegate
            {
                pictureBoxCUV3.BackColor = (cellStatus3 & (ushort)MemoryBms.CellStatus.Cell_Status_CUV_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CUV
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCUV4.Invoke((ThreadStart)delegate
            {
                pictureBoxCUV4.BackColor = (cellStatus4 & (ushort)MemoryBms.CellStatus.Cell_Status_CUV_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CUV
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCVMAX1.Invoke((ThreadStart)delegate
            {
                pictureBoxCVMAX1.BackColor = (cellStatus1 & (ushort)MemoryBms.CellStatus.Cell_Status_CVMAX_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CVMAX
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCVMAX2.Invoke((ThreadStart)delegate
            {
                pictureBoxCVMAX2.BackColor = (cellStatus2 & (ushort)MemoryBms.CellStatus.Cell_Status_CVMAX_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CVMAX
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCVMAX3.Invoke((ThreadStart)delegate
            {
                pictureBoxCVMAX3.BackColor = (cellStatus3 & (ushort)MemoryBms.CellStatus.Cell_Status_CVMAX_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CVMAX
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCVMAX4.Invoke((ThreadStart)delegate
            {
                pictureBoxCVMAX4.BackColor = (cellStatus4 & (ushort)MemoryBms.CellStatus.Cell_Status_CVMAX_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CVMAX
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCVMIN1.Invoke((ThreadStart)delegate
            {
                pictureBoxCVMIN1.BackColor = (cellStatus1 & (ushort)MemoryBms.CellStatus.Cell_Status_CVMIN_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CVMIN
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCVMIN2.Invoke((ThreadStart)delegate
            {
                pictureBoxCVMIN2.BackColor = (cellStatus2 & (ushort)MemoryBms.CellStatus.Cell_Status_CVMIN_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CVMIN
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCVMIN3.Invoke((ThreadStart)delegate
            {
                pictureBoxCVMIN3.BackColor = (cellStatus3 & (ushort)MemoryBms.CellStatus.Cell_Status_CVMIN_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CVMIN
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCVMIN4.Invoke((ThreadStart)delegate
            {
                pictureBoxCVMIN4.BackColor = (cellStatus4 & (ushort)MemoryBms.CellStatus.Cell_Status_CVMIN_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_CVMIN
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxBAL1.Invoke((ThreadStart)delegate
            {
                pictureBoxBAL1.BackColor = (cellStatus1 & (ushort)MemoryBms.CellStatus.Cell_Status_Balancing_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_Balancing
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxBAL2.Invoke((ThreadStart)delegate
            {
                pictureBoxBAL2.BackColor = (cellStatus2 & (ushort)MemoryBms.CellStatus.Cell_Status_Balancing_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_Balancing
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxBAL3.Invoke((ThreadStart)delegate
            {
                pictureBoxBAL3.BackColor = (cellStatus3 & (ushort)MemoryBms.CellStatus.Cell_Status_Balancing_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_Balancing
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxBAL4.Invoke((ThreadStart)delegate
            {
                pictureBoxBAL4.BackColor = (cellStatus4 & (ushort)MemoryBms.CellStatus.Cell_Status_Balancing_Msk) == (ushort)MemoryBms.CellStatus.Cell_Status_Balancing
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
        }

        private void UpdateCellSOC()
        {
            ushort chargeStatus = _batteryModule.GetMemory().StatusRegisters.chargeStatus;
            pictureBoxSOC20.Invoke((ThreadStart)delegate
            {
                pictureBoxSOC20.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC20_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC20
                    ? Color.OrangeRed
                    : Color.Gainsboro;
            });
            pictureBoxSOC40.Invoke((ThreadStart)delegate
            {
                pictureBoxSOC40.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC40_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC40
                    ? Color.DarkOrange
                    : Color.Gainsboro;
            });
            pictureBoxSOC60.Invoke((ThreadStart)delegate
            {
                pictureBoxSOC60.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC60_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC60
                    ? Color.LimeGreen
                    : Color.Gainsboro;
            });
            pictureBoxSOC80.Invoke((ThreadStart)delegate
            {
                pictureBoxSOC80.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC80_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC80
                    ? Color.LimeGreen
                    : Color.Gainsboro;
            });
            pictureBoxSOC100.Invoke((ThreadStart)delegate
            {
                pictureBoxSOC100.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC100_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC100
                    ? Color.LimeGreen
                    : Color.Gainsboro;
            });
        }

        private void UpdateCellChargeAlert()
        {
            ushort chargeAlert = _batteryModule.GetMemory().StatusRegisters.chargeAlert;
            pictureBoxCharge_Status_CC_A.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_CC_A.BackColor = (chargeAlert & (ushort)MemoryBms.ChargeStatus.Charge_Status_CC_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_CC
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_DC_A.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_DC_A.BackColor = (chargeAlert & (ushort)MemoryBms.ChargeStatus.Charge_Status_DC_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_DC
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_RCA_A.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_RCA_A.BackColor = (chargeAlert & (ushort)MemoryBms.ChargeStatus.Charge_Status_RCA_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_RCA
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_BAL_A.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_BAL_A.BackColor = (chargeAlert & (ushort)MemoryBms.ChargeStatus.Charge_Status_BAL_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_BAL
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_SOC20_A.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_SOC20_A.BackColor = (chargeAlert & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC20_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC20
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_SOC40_A.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_SOC40_A.BackColor = (chargeAlert & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC40_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC40
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_SOC60_A.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_SOC60_A.BackColor = (chargeAlert & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC60_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC60
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_SOC80_A.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_SOC80_A.BackColor = (chargeAlert & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC80_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC80
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_SOC100_A.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_SOC100_A.BackColor = (chargeAlert & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC100_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC100
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
        }

        private void UpdateCellChargeStatus()
        {
            ushort chargeStatus = _batteryModule.GetMemory().StatusRegisters.chargeStatus;
            pictureBoxCharge_Status_CC_S.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_CC_S.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_CC_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_CC
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_DC_S.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_DC_S.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_DC_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_DC
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_RCA_S.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_RCA_S.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_RCA_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_RCA
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_BAL_S.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_BAL_S.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_BAL_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_BAL
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_SOC20_S.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_SOC20_S.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC20_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC20
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_SOC40_S.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_SOC40_S.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC40_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC40
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_SOC60_S.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_SOC60_S.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC60_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC60
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_SOC80_S.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_SOC80_S.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC80_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC80
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxCharge_Status_SOC100_S.Invoke((ThreadStart)delegate
            {
                pictureBoxCharge_Status_SOC100_S.BackColor = (chargeStatus & (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC100_Msk) == (ushort)MemoryBms.ChargeStatus.Charge_Status_SOC100
                    ? pictureBoxGreen.BackColor
                    : Color.Gainsboro;
            });
        }

        private void UpdateCellSafetyAlert()
        {
            ushort safetyAlert = _batteryModule.GetMemory().StatusRegisters.safetyAlert;
            pictureBoxSafety_Status_COVC_A.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_COVC_A.BackColor = (safetyAlert & (ushort)MemoryBms.SafetyStatus.Safety_Status_COVC_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_COVC
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_COVT_A.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_COVT_A.BackColor = (safetyAlert & (ushort)MemoryBms.SafetyStatus.Safety_Status_COVT_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_COVT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_CUV_A.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_CUV_A.BackColor = (safetyAlert & (ushort)MemoryBms.SafetyStatus.Safety_Status_CUV_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_CUV
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_COT_A.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_COT_A.BackColor = (safetyAlert & (ushort)MemoryBms.SafetyStatus.Safety_Status_COT_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_COT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_CUT_A.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_CUT_A.BackColor = (safetyAlert & (ushort)MemoryBms.SafetyStatus.Safety_Status_CUT_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_CUT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_UTC_A.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_UTC_A.BackColor = (safetyAlert & (ushort)MemoryBms.SafetyStatus.Safety_Status_UTC_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_UTC
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_UTD_A.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_UTD_A.BackColor = (safetyAlert & (ushort)MemoryBms.SafetyStatus.Safety_Status_UTD_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_UTD
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_ROT_A.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_ROT_A.BackColor = (safetyAlert & (ushort)MemoryBms.SafetyStatus.Safety_Status_ROT_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_ROT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_RUT_A.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_RUT_A.BackColor = (safetyAlert & (ushort)MemoryBms.SafetyStatus.Safety_Status_RUT_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_RUT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
        }

        private void UpdateCellSafetyStatus()
        {
            ushort safetyStatus = _batteryModule.GetMemory().StatusRegisters.safetyStatus;
            pictureBoxSafety_Status_COVC_S.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_COVC_S.BackColor = (safetyStatus & (ushort)MemoryBms.SafetyStatus.Safety_Status_COVC_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_COVC
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_COVT_S.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_COVT_S.BackColor = (safetyStatus & (ushort)MemoryBms.SafetyStatus.Safety_Status_COVT_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_COVT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_CUV_S.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_CUV_S.BackColor = (safetyStatus & (ushort)MemoryBms.SafetyStatus.Safety_Status_CUV_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_CUV
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_COT_S.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_COT_S.BackColor = (safetyStatus & (ushort)MemoryBms.SafetyStatus.Safety_Status_COT_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_COT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_CUT_S.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_CUT_S.BackColor = (safetyStatus & (ushort)MemoryBms.SafetyStatus.Safety_Status_CUT_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_CUT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_UTC_S.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_UTC_S.BackColor = (safetyStatus & (ushort)MemoryBms.SafetyStatus.Safety_Status_UTC_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_UTC
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_UTD_S.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_UTD_S.BackColor = (safetyStatus & (ushort)MemoryBms.SafetyStatus.Safety_Status_UTD_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_UTD
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_ROT_S.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_ROT_S.BackColor = (safetyStatus & (ushort)MemoryBms.SafetyStatus.Safety_Status_ROT_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_ROT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
            pictureBoxSafety_Status_RUT_S.Invoke((ThreadStart)delegate
            {
                pictureBoxSafety_Status_RUT_S.BackColor = (safetyStatus & (ushort)MemoryBms.SafetyStatus.Safety_Status_RUT_Msk) == (ushort)MemoryBms.SafetyStatus.Safety_Status_RUT
                    ? pictureBoxRed.BackColor
                    : Color.Gainsboro;
            });
        }

        private void UpdateBatteryStatus()
        {
            ushort batteryStatus = _batteryModule.GetMemory().StatusRegisters.batteryStatus;
            pictureBoxBattery_Status_TDA.Invoke((ThreadStart)delegate
            {
                pictureBoxBattery_Status_TDA.BackColor = (batteryStatus & (ushort)MemoryBms.BatteryStatus.Battery_Status_TDA_Msk) == (ushort)MemoryBms.BatteryStatus.Battery_Status_TDA
                    ? Color.DarkOrange
                    : Color.Gainsboro;
            });
            pictureBoxBattery_Status_TCA.Invoke((ThreadStart)delegate
            {
                pictureBoxBattery_Status_TCA.BackColor = (batteryStatus & (ushort)MemoryBms.BatteryStatus.Battery_Status_TCA_Msk) == (ushort)MemoryBms.BatteryStatus.Battery_Status_TCA
                    ? Color.DarkOrange
                    : Color.Gainsboro;
            });
            pictureBoxBattery_Status_FC.Invoke((ThreadStart)delegate
            {
                pictureBoxBattery_Status_FC.BackColor = (batteryStatus & (ushort)MemoryBms.BatteryStatus.Battery_Status_FC_Msk) == (ushort)MemoryBms.BatteryStatus.Battery_Status_FC
                    ? Color.DarkOrange
                    : Color.Gainsboro;
            });
            pictureBoxBattery_Status_FD.Invoke((ThreadStart)delegate
            {
                pictureBoxBattery_Status_FD.BackColor = (batteryStatus & (ushort)MemoryBms.BatteryStatus.Battery_Status_FD_Msk) == (ushort)MemoryBms.BatteryStatus.Battery_Status_FD
                    ? Color.DarkOrange
                    : Color.Gainsboro;
            });
            pictureBoxBattery_Status_RCA.Invoke((ThreadStart)delegate
            {
                pictureBoxBattery_Status_RCA.BackColor = (batteryStatus & (ushort)MemoryBms.BatteryStatus.Battery_Status_RCA_Msk) == (ushort)MemoryBms.BatteryStatus.Battery_Status_RCA
                    ? Color.DarkOrange
                    : Color.Gainsboro;
            });
            pictureBoxBattery_Status_OCA.Invoke((ThreadStart)delegate
            {
                pictureBoxBattery_Status_OCA.BackColor = (batteryStatus & (ushort)MemoryBms.BatteryStatus.Battery_Status_OCA_Msk) == (ushort)MemoryBms.BatteryStatus.Battery_Status_OCA
                    ? Color.DarkOrange
                    : Color.Gainsboro;
            });
            pictureBoxBattery_Status_OTA.Invoke((ThreadStart)delegate
            {
                pictureBoxBattery_Status_OTA.BackColor = (batteryStatus & (ushort)MemoryBms.BatteryStatus.Battery_Status_OTA_Msk) == (ushort)MemoryBms.BatteryStatus.Battery_Status_OTA
                    ? Color.DarkOrange
                    : Color.Gainsboro;
            });
            pictureBoxBattery_Status_PS.Invoke((ThreadStart)delegate
            {
                pictureBoxBattery_Status_PS.BackColor = (batteryStatus & (ushort)MemoryBms.BatteryStatus.Battery_Status_PS_Msk) == (ushort)MemoryBms.BatteryStatus.Battery_Status_PS
                    ? Color.DarkOrange
                    : Color.Gainsboro;
            });
        }

        public void UpdateSettingsCharge()
        {
            textBoxBalance_Voltage_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxBalance_Voltage_Threshold.Text = _batteryModule.GetMemory().SettingsCharge.Balance_Voltage_Threshold.ToString();
            });
            textBoxBalance_Voltage_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxBalance_Voltage_Recovery.Text = _batteryModule.GetMemory().SettingsCharge.Balance_Voltage_Recovery.ToString();
            });
            textBoxBalance_Time.Invoke((ThreadStart)delegate
            {
                textBoxBalance_Time.Text = _batteryModule.GetMemory().SettingsCharge.Balance_Time.ToString();
            });
            textBoxBalance_Delta_Voltage.Invoke((ThreadStart)delegate
            {
                textBoxBalance_Delta_Voltage.Text = _batteryModule.GetMemory().SettingsCharge.Balance_Delta_Voltage.ToString();
            });
            textBoxCharge_Voltage_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxCharge_Voltage_Threshold.Text = _batteryModule.GetMemory().SettingsCharge.Charge_Voltage_Threshold.ToString();
            });
            textBoxCharge_Voltage_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCharge_Voltage_Recovery.Text = _batteryModule.GetMemory().SettingsCharge.Charge_Voltage_Recovery.ToString();
            });
            textBoxCharge_Completion_Time.Invoke((ThreadStart)delegate
            {
                textBoxCharge_Completion_Time.Text = _batteryModule.GetMemory().SettingsCharge.Charge_Completion_Time.ToString();
            });
            textBoxDischarge_Voltage_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxDischarge_Voltage_Recovery.Text = _batteryModule.GetMemory().SettingsCharge.Discharge_Voltage_Recovery.ToString();
            });
            textBoxDischarge_Completion_Time.Invoke((ThreadStart)delegate
            {
                textBoxDischarge_Completion_Time.Text = _batteryModule.GetMemory().SettingsCharge.Discharge_Completion_Time.ToString();
            });
            textBoxRemaining_Capacity_Alarm_Percent.Invoke((ThreadStart)delegate
            {
                textBoxRemaining_Capacity_Alarm_Percent.Text = _batteryModule.GetMemory().SettingsCharge.Remaining_Capacity_Alarm_Percent.ToString();
            });
            textBoxCV20_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxCV20_Threshold.Text = _batteryModule.GetMemory().SettingsCharge.CV20_Threshold.ToString();
            });
            textBoxCV20_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCV20_Recovery.Text = _batteryModule.GetMemory().SettingsCharge.CV20_Recovery.ToString();
            });
            textBoxCV20_Time.Invoke((ThreadStart)delegate
            {
                textBoxCV20_Time.Text = _batteryModule.GetMemory().SettingsCharge.CV20_Time.ToString();
            });
            textBoxCV40_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxCV40_Threshold.Text = _batteryModule.GetMemory().SettingsCharge.CV40_Threshold.ToString();
            });
            textBoxCV40_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCV40_Recovery.Text = _batteryModule.GetMemory().SettingsCharge.CV40_Recovery.ToString();
            });
            textBoxCV40_Time.Invoke((ThreadStart)delegate
            {
                textBoxCV40_Time.Text = _batteryModule.GetMemory().SettingsCharge.CV40_Time.ToString();
            });
            textBoxCV60_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxCV60_Threshold.Text = _batteryModule.GetMemory().SettingsCharge.CV60_Threshold.ToString();
            });
            textBoxCV60_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCV60_Recovery.Text = _batteryModule.GetMemory().SettingsCharge.CV60_Recovery.ToString();
            });
            textBoxCV60_Time.Invoke((ThreadStart)delegate
            {
                textBoxCV60_Time.Text = _batteryModule.GetMemory().SettingsCharge.CV60_Time.ToString();
            });
            textBoxCV80_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxCV80_Threshold.Text = _batteryModule.GetMemory().SettingsCharge.CV80_Threshold.ToString();
            });
            textBoxCV80_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCV80_Recovery.Text = _batteryModule.GetMemory().SettingsCharge.CV80_Recovery.ToString();
            });
            textBoxCV80_Time.Invoke((ThreadStart)delegate
            {
                textBoxCV80_Time.Text = _batteryModule.GetMemory().SettingsCharge.CV80_Time.ToString();
            });
            textBoxCV100_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxCV100_Threshold.Text = _batteryModule.GetMemory().SettingsCharge.CV100_Threshold.ToString();
            });
            textBoxCV100_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCV100_Recovery.Text = _batteryModule.GetMemory().SettingsCharge.CV100_Recovery.ToString();
            });
            textBoxCV100_Time.Invoke((ThreadStart)delegate
            {
                textBoxCV100_Time.Text = _batteryModule.GetMemory().SettingsCharge.CV100_Time.ToString();
            });
        }

        public void UpdateSettingsAlarm()
        {
            textBoxCOVC_Threshold.Invoke((ThreadStart) delegate
            {
                textBoxCOVC_Threshold.Text = _batteryModule.GetMemory().SettingsAlarm.COVC_Threshold.ToString();
            });
            textBoxCOVC_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCOVC_Recovery.Text = _batteryModule.GetMemory().SettingsAlarm.COVC_Recovery.ToString();
            });
            textBoxCOVC_Time.Invoke((ThreadStart)delegate
            {
                textBoxCOVC_Time.Text = _batteryModule.GetMemory().SettingsAlarm.COVC_Time.ToString();
            });
            textBoxCOVT_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxCOVT_Threshold.Text = _batteryModule.GetMemory().SettingsAlarm.COVT_Threshold.ToString();
            });
            textBoxCOVT_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCOVT_Recovery.Text = _batteryModule.GetMemory().SettingsAlarm.COVT_Recovery.ToString();
            });
            textBoxCOVT_Time.Invoke((ThreadStart)delegate
            {
                textBoxCOVT_Time.Text = _batteryModule.GetMemory().SettingsAlarm.COVT_Time.ToString();
            });
            textBoxCUV_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxCUV_Threshold.Text = _batteryModule.GetMemory().SettingsAlarm.CUV_Threshold.ToString();
            });
            textBoxCUV_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCUV_Recovery.Text = _batteryModule.GetMemory().SettingsAlarm.CUV_Recovery.ToString();
            });
            textBoxCUV_Time.Invoke((ThreadStart)delegate
            {
                textBoxCUV_Time.Text = _batteryModule.GetMemory().SettingsAlarm.CUV_Time.ToString();
            });
            textBoxCOT_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxCOT_Threshold.Text = _batteryModule.GetMemory().SettingsAlarm.COT_Threshold.ToString();
            });
            textBoxCOT_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCOT_Recovery.Text = _batteryModule.GetMemory().SettingsAlarm.COT_Recovery.ToString();
            });
            textBoxCOT_Time.Invoke((ThreadStart)delegate
            {
                textBoxCOT_Time.Text = _batteryModule.GetMemory().SettingsAlarm.COT_Time.ToString();
            });
            textBoxCUT_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxCUT_Threshold.Text = _batteryModule.GetMemory().SettingsAlarm.CUT_Threshold.ToString();
            });
            textBoxCUT_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxCUT_Recovery.Text = _batteryModule.GetMemory().SettingsAlarm.CUT_Recovery.ToString();
            });
            textBoxCUT_Time.Invoke((ThreadStart)delegate
            {
                textBoxCUT_Time.Text = _batteryModule.GetMemory().SettingsAlarm.CUT_Time.ToString();
            });
            textBoxROT_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxROT_Threshold.Text = _batteryModule.GetMemory().SettingsAlarm.ROT_Threshold.ToString();
            });
            textBoxROT_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxROT_Recovery.Text = _batteryModule.GetMemory().SettingsAlarm.ROT_Recovery.ToString();
            });
            textBoxROT_Time.Invoke((ThreadStart)delegate
            {
                textBoxROT_Time.Text = _batteryModule.GetMemory().SettingsAlarm.ROT_Time.ToString();
            });
            textBoxRUT_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxRUT_Threshold.Text = _batteryModule.GetMemory().SettingsAlarm.RUT_Threshold.ToString();
            });
            textBoxRUT_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxRUT_Recovery.Text = _batteryModule.GetMemory().SettingsAlarm.RUT_Recovery.ToString();
            });
            textBoxRUT_Time.Invoke((ThreadStart)delegate
            {
                textBoxRUT_Time.Text = _batteryModule.GetMemory().SettingsAlarm.RUT_Time.ToString();
            });
            textBoxUTC_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxUTC_Threshold.Text = _batteryModule.GetMemory().SettingsAlarm.UTC_Threshold.ToString();
            });
            textBoxUTC_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxUTC_Recovery.Text = _batteryModule.GetMemory().SettingsAlarm.UTC_Recovery.ToString();
            });
            textBoxUTC_Time.Invoke((ThreadStart)delegate
            {
                textBoxUTC_Time.Text = _batteryModule.GetMemory().SettingsAlarm.UTC_Time.ToString();
            });
            textBoxUTD_Threshold.Invoke((ThreadStart)delegate
            {
                textBoxUTD_Threshold.Text = _batteryModule.GetMemory().SettingsAlarm.UTD_Threshold.ToString();
            });
            textBoxUTD_Recovery.Invoke((ThreadStart)delegate
            {
                textBoxUTD_Recovery.Text = _batteryModule.GetMemory().SettingsAlarm.UTD_Recovery.ToString();
            });
            textBoxUTD_Time.Invoke((ThreadStart)delegate
            {
                textBoxUTD_Time.Text = _batteryModule.GetMemory().SettingsAlarm.UTD_Time.ToString();
            });
        }

        private void buttonClearMain_Click(object sender, EventArgs e)
        {
            textBoxReceiveMain.Text = string.Empty;
            textBoxTransmitMain.Text = string.Empty;
            ClearMeasurementsFields();
            ClearRegisters();
        }

        private void ClearMeasurementsFields()
        {
            textBoxBatteryVoltage.Text = string.Empty;
            textBoxCellVoltageMax.Text = string.Empty;
            textBoxCellVoltageMin.Text = string.Empty;
            textBoxRadiatorTemperature.Text = string.Empty;
            textBoxSOC.Text = string.Empty;
            textBoxСellVoltage1.Text = string.Empty;
            textBoxСellVoltage2.Text = string.Empty;
            textBoxСellVoltage3.Text = string.Empty;
            textBoxСellVoltage4.Text = string.Empty;
            textBoxADCVoltage1.Text = string.Empty;
            textBoxADCVoltage2.Text = string.Empty;
            textBoxADCVoltage3.Text = string.Empty;
            textBoxADCVoltage4.Text = string.Empty;
            pictureBoxCOVC1.BackColor = Color.Gainsboro;
            pictureBoxCOVC2.BackColor = Color.Gainsboro;
            pictureBoxCOVC3.BackColor = Color.Gainsboro;
            pictureBoxCOVC4.BackColor = Color.Gainsboro;
            pictureBoxCOVT1.BackColor = Color.Gainsboro;
            pictureBoxCOVT2.BackColor = Color.Gainsboro;
            pictureBoxCOVT3.BackColor = Color.Gainsboro;
            pictureBoxCOVT4.BackColor = Color.Gainsboro;
            pictureBoxCUV1.BackColor = Color.Gainsboro;
            pictureBoxCUV2.BackColor = Color.Gainsboro;
            pictureBoxCUV3.BackColor = Color.Gainsboro;
            pictureBoxCUV4.BackColor = Color.Gainsboro;
            pictureBoxCVMAX1.BackColor = Color.Gainsboro;
            pictureBoxCVMAX2.BackColor = Color.Gainsboro;
            pictureBoxCVMAX3.BackColor = Color.Gainsboro;
            pictureBoxCVMAX4.BackColor = Color.Gainsboro;
            pictureBoxCVMIN1.BackColor = Color.Gainsboro;
            pictureBoxCVMIN2.BackColor = Color.Gainsboro;
            pictureBoxCVMIN3.BackColor = Color.Gainsboro;
            pictureBoxCVMIN4.BackColor = Color.Gainsboro;
            pictureBoxBAL1.BackColor = Color.Gainsboro;
            pictureBoxBAL2.BackColor = Color.Gainsboro;
            pictureBoxBAL3.BackColor = Color.Gainsboro;
            pictureBoxBAL4.BackColor = Color.Gainsboro;
            pictureBoxSOC100.BackColor = Color.Gainsboro;
            pictureBoxSOC80.BackColor = Color.Gainsboro;
            pictureBoxSOC60.BackColor = Color.Gainsboro;
            pictureBoxSOC40.BackColor = Color.Gainsboro;
            pictureBoxSOC20.BackColor = Color.Gainsboro;
            pictureBoxRT_status.BackColor = Color.Gainsboro;

            progressBarСellVoltage1.Invoke((ThreadStart)delegate
            {
                progressBarСellVoltage1.Value = progressBarСellVoltage1.Minimum;
            });
            progressBarСellVoltage2.Invoke((ThreadStart)delegate
            {
                progressBarСellVoltage2.Value = progressBarСellVoltage2.Minimum;
            });
            progressBarСellVoltage3.Invoke((ThreadStart)delegate
            {
                progressBarСellVoltage3.Value = progressBarСellVoltage3.Minimum;
            });
            progressBarСellVoltage4.Invoke((ThreadStart)delegate
            {
                progressBarСellVoltage4.Value = progressBarСellVoltage4.Minimum;
            });
        }

        private void ClearRegisters()
        {
            ClearCellChargeAlert();
            ClearCellChargeStatus();
            ClearCellSafetyAlert();
            ClearCellSafetyStatus();
            ClearBatteryStatus();
        }

        private void ClearCellChargeAlert()
        {
            pictureBoxCharge_Status_CC_A.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_DC_A.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_RCA_A.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_BAL_A.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_SOC20_A.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_SOC40_A.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_SOC60_A.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_SOC80_A.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_SOC100_A.BackColor = Color.Gainsboro;
        }

        private void ClearCellChargeStatus()
        {
            pictureBoxCharge_Status_CC_S.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_DC_S.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_RCA_S.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_BAL_S.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_SOC20_S.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_SOC40_S.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_SOC60_S.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_SOC80_S.BackColor = Color.Gainsboro;
            pictureBoxCharge_Status_SOC100_S.BackColor = Color.Gainsboro;
        }

        private void ClearCellSafetyAlert()
        {
            pictureBoxSafety_Status_COVC_A.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_COVT_A.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_CUV_A.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_COT_A.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_CUT_A.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_UTC_A.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_UTD_A.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_ROT_A.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_RUT_A.BackColor = Color.Gainsboro;
        }

        private void ClearCellSafetyStatus()
        {
            pictureBoxSafety_Status_COVC_S.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_COVT_S.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_CUV_S.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_COT_S.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_CUT_S.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_UTC_S.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_UTD_S.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_ROT_S.BackColor = Color.Gainsboro;
            pictureBoxSafety_Status_RUT_S.BackColor = Color.Gainsboro;
        }

        private void ClearBatteryStatus()
        {
            pictureBoxBattery_Status_TDA.BackColor = Color.Gainsboro;
            pictureBoxBattery_Status_TCA.BackColor = Color.Gainsboro;
            pictureBoxBattery_Status_FC.BackColor = Color.Gainsboro;
            pictureBoxBattery_Status_FD.BackColor = Color.Gainsboro;
            pictureBoxBattery_Status_RCA.BackColor = Color.Gainsboro;
            pictureBoxBattery_Status_OCA.BackColor = Color.Gainsboro;
            pictureBoxBattery_Status_OTA.BackColor = Color.Gainsboro;
            pictureBoxBattery_Status_PS.BackColor = Color.Gainsboro;
        }

        private void buttonClearDebug_Click(object sender, EventArgs e)
        {
            textBoxReceiveDebug.Text = string.Empty;
            textBoxTransmitDebug.Text = string.Empty;
        }

        private void buttonClearSettingsCalibration_Click(object sender, EventArgs e)
        {
            textBoxReceiveSettingsCalibration.Text = string.Empty;
            textBoxTransmitSettingsCalibration.Text = string.Empty;
            ClearSettingsCalibrationFields();
        }

        private void ClearSettingsCalibrationFields()
        {
            textBoxUADC1Calibration.Text = string.Empty;
            textBoxUADC2Calibration.Text = string.Empty;
            textBoxUADC3Calibration.Text = string.Empty;
            textBoxUADC4Calibration.Text = string.Empty;
            textBoxADCTempCalibration.Text = string.Empty;
            textBoxADC_U1_K.Text = string.Empty;
            textBoxADC_U2_K.Text = string.Empty;
            textBoxADC_U3_K.Text = string.Empty;
            textBoxADC_U4_K.Text = string.Empty;
            textBoxADC_TEMP_K.Text = string.Empty;
        }

        private void buttonClearSettingsCharge_Click(object sender, EventArgs e)
        {
            textBoxReceiveSettingsCharge.Text = string.Empty;
            textBoxTransmitSettingsCharge.Text = string.Empty;
            ClearSettingsChargeFields();
        }

        private void ClearSettingsChargeFields()
        {
            textBoxBalance_Voltage_Threshold.Text = string.Empty;
            textBoxBalance_Voltage_Recovery.Text = string.Empty;
            textBoxBalance_Time.Text = string.Empty;
            textBoxBalance_Delta_Voltage.Text = string.Empty;
            textBoxCharge_Voltage_Threshold.Text = string.Empty;
            textBoxCharge_Voltage_Recovery.Text = string.Empty;
            textBoxCharge_Completion_Time.Text = string.Empty;
            textBoxDischarge_Voltage_Recovery.Text = string.Empty;
            textBoxDischarge_Completion_Time.Text = string.Empty;
            textBoxRemaining_Capacity_Alarm_Percent.Text = string.Empty;
            textBoxCV20_Threshold.Text = string.Empty;
            textBoxCV20_Recovery.Text = string.Empty;
            textBoxCV20_Time.Text = string.Empty;
            textBoxCV40_Threshold.Text = string.Empty;
            textBoxCV40_Recovery.Text = string.Empty;
            textBoxCV40_Time.Text = string.Empty;
            textBoxCV60_Threshold.Text = string.Empty;
            textBoxCV60_Recovery.Text = string.Empty;
            textBoxCV60_Time.Text = string.Empty;
            textBoxCV80_Threshold.Text = string.Empty;
            textBoxCV80_Recovery.Text = string.Empty;
            textBoxCV80_Time.Text = string.Empty;
            textBoxCV100_Threshold.Text = string.Empty;
            textBoxCV100_Recovery.Text = string.Empty;
            textBoxCV100_Time.Text = string.Empty;
        }

        private void buttonClearSettingsAlarm_Click(object sender, EventArgs e)
        {
            textBoxReceiveSettingsAlarm.Text = string.Empty;
            textBoxTransmitSettingsAlarm.Text = string.Empty;
            ClearSettingsAlarmFields();
        }

        private void ClearSettingsAlarmFields()
        {
            textBoxCOVC_Threshold.Text = string.Empty;
            textBoxCOVC_Recovery.Text = string.Empty;
            textBoxCOVC_Time.Text = string.Empty;
            textBoxCOVT_Threshold.Text = string.Empty;
            textBoxCOVT_Recovery.Text = string.Empty;
            textBoxCOVT_Time.Text = string.Empty;
            textBoxCUV_Threshold.Text = string.Empty;
            textBoxCUV_Recovery.Text = string.Empty;
            textBoxCUV_Time.Text = string.Empty;
            textBoxCOT_Threshold.Text = string.Empty;
            textBoxCOT_Recovery.Text = string.Empty;
            textBoxCOT_Time.Text = string.Empty;
            textBoxCUT_Threshold.Text = string.Empty;
            textBoxCUT_Recovery.Text = string.Empty;
            textBoxCUT_Time.Text = string.Empty;
            textBoxROT_Threshold.Text = string.Empty;
            textBoxROT_Recovery.Text = string.Empty;
            textBoxROT_Time.Text = string.Empty;
            textBoxRUT_Threshold.Text = string.Empty;
            textBoxRUT_Recovery.Text = string.Empty;
            textBoxRUT_Time.Text = string.Empty;
            textBoxUTC_Threshold.Text = string.Empty;
            textBoxUTC_Recovery.Text = string.Empty;
            textBoxUTC_Time.Text = string.Empty;
            textBoxUTD_Threshold.Text = string.Empty;
            textBoxUTD_Recovery.Text = string.Empty;
            textBoxUTD_Time.Text = string.Empty;
        }

        private void buttonDebugEnable_Click(object sender, EventArgs e)
        {
            buttonDebugDisable.Enabled = true;
            buttonDebugEnable.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcEnableDebugging);
            _batteryModule.ResetLastCommand();
        }

        private void buttonDebugDisable_Click(object sender, EventArgs e)
        {
            buttonDebugEnable.Enabled = true;
            buttonDebugDisable.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcDisableDebugging);
            _batteryModule.ResetLastCommand();
        }

        private void buttonContactorOn_Click(object sender, EventArgs e)
        {
            buttonContactorOff.Enabled = true;
            buttonContactorOn.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcContactorEnable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonContactorOff_Click(object sender, EventArgs e)
        {
            buttonContactorOn.Enabled = true;
            buttonContactorOff.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcContactorDisable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonChargerOn_Click(object sender, EventArgs e)
        {
            buttonChargerOff.Enabled = true;
            buttonChargerOn.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcChargerEnable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonChargerOff_Click(object sender, EventArgs e)
        {
            buttonChargerOn.Enabled = true;
            buttonChargerOff.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcChargerDisable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonMotorOn_Click(object sender, EventArgs e)
        {
            buttonMotorOff.Enabled = true;
            buttonMotorOn.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcMotorEnable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonMotorOff_Click(object sender, EventArgs e)
        {
            buttonMotorOn.Enabled = true;
            buttonMotorOff.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcMotorDisable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonBalance1On_Click(object sender, EventArgs e)
        {
            buttonBalance1Off.Enabled = true;
            buttonBalance1On.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcBalancingCell1Enable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonBalance1Off_Click(object sender, EventArgs e)
        {
            buttonBalance1On.Enabled = true;
            buttonBalance1Off.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcBalancingCell1Disable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonBalance2On_Click(object sender, EventArgs e)
        {
            buttonBalance2Off.Enabled = true;
            buttonBalance2On.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcBalancingCell2Enable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonBalance2Off_Click(object sender, EventArgs e)
        {
            buttonBalance2On.Enabled = true;
            buttonBalance2Off.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcBalancingCell2Disable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonBalance3On_Click(object sender, EventArgs e)
        {
            buttonBalance3Off.Enabled = true;
            buttonBalance3On.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcBalancingCell3Enable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonBalance3Off_Click(object sender, EventArgs e)
        {
            buttonBalance3On.Enabled = true;
            buttonBalance3Off.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcBalancingCell3Disable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonBalance4On_Click(object sender, EventArgs e)
        {
            buttonBalance4Off.Enabled = true;
            buttonBalance4On.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcBalancingCell4Enable);
            _batteryModule.ResetLastCommand();
        }

        private void buttonBalance4Off_Click(object sender, EventArgs e)
        {
            buttonBalance4On.Enabled = true;
            buttonBalance4Off.Enabled = false;
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcBalancingCell4Disable);
            _batteryModule.ResetLastCommand();
        }

        public void UpdateTextInCalibrationK(MemoryBms.CellIndex index)
        {
            switch (index)
            {
                case MemoryBms.CellIndex.Cell_1:
                    textBoxADC_U1_K.Invoke((ThreadStart)delegate
                    {
                        textBoxADC_U1_K.Text = _batteryModule.GetCalibrationCoefficient().ToString();
                    });
                    break;
                case MemoryBms.CellIndex.Cell_2:
                    textBoxADC_U2_K.Invoke((ThreadStart)delegate
                    {
                        textBoxADC_U2_K.Text = _batteryModule.GetCalibrationCoefficient().ToString();
                    });
                    break;
                case MemoryBms.CellIndex.Cell_3:
                    textBoxADC_U3_K.Invoke((ThreadStart)delegate
                    {
                        textBoxADC_U3_K.Text = _batteryModule.GetCalibrationCoefficient().ToString();
                    });
                    break;
                case MemoryBms.CellIndex.Cell_4:
                    textBoxADC_U4_K.Invoke((ThreadStart)delegate
                    {
                        textBoxADC_U4_K.Text = _batteryModule.GetCalibrationCoefficient().ToString();
                    });
                    break;
                case MemoryBms.CellIndex.Temperature:
                    textBoxADC_TEMP_K.Invoke((ThreadStart)delegate
                    {
                        textBoxADC_TEMP_K.Text = _batteryModule.GetCalibrationCoefficient().ToString();
                    });
                    break;
            }
        }

        private void buttonCalibrationUADC1_Click(object sender, EventArgs e)
        {
            _batteryModule.SetCalibrationValue((ushort)int.Parse(textBoxUADC1Calibration.Text));
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcCalibrationU1);
        }

        private void buttonCalibrationUADC2_Click(object sender, EventArgs e)
        {
            _batteryModule.SetCalibrationValue((ushort)int.Parse(textBoxUADC2Calibration.Text));
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcCalibrationU2);
        }

        private void buttonCalibrationUADC3_Click(object sender, EventArgs e)
        {
            _batteryModule.SetCalibrationValue((ushort)int.Parse(textBoxUADC3Calibration.Text));
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcCalibrationU3);
        }

        private void buttonCalibrationUADC4_Click(object sender, EventArgs e)
        {
            _batteryModule.SetCalibrationValue((ushort)int.Parse(textBoxUADC4Calibration.Text));
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcCalibrationU4);
        }

        private void buttonCalibrationTempR_Click(object sender, EventArgs e)
        {
            _batteryModule.SetCalibrationValue((ushort)int.Parse(textBoxADCTempCalibration.Text));
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcCalibrationTemp);
        }

        private void buttonSaveSettings_Click(object sender, EventArgs e)
        {
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcSaveSettings);
            _batteryModule.ResetLastCommand();
        }

        private void buttonGetSettingsCharge_Click(object sender, EventArgs e)
        {
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcSendSettingsCharge);
        }

        private void buttonGetSettingsAlarm_Click(object sender, EventArgs e)
        {
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcSendSettingsAlarm);
        }

        private void buttonSetSettingsCharge_Click(object sender, EventArgs e)
        {
            _batteryModule.GetMemory().SetSettingsCharge_Balance_Voltage_Threshold((ushort)int.Parse(textBoxBalance_Voltage_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsCharge_Balance_Voltage_Recovery((ushort)int.Parse(textBoxBalance_Voltage_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsCharge_Balance_Time((byte)int.Parse(textBoxBalance_Time.Text));
            _batteryModule.GetMemory().SetSettingsCharge_Balance_Delta_Voltage((byte)int.Parse(textBoxBalance_Delta_Voltage.Text));
            _batteryModule.GetMemory().SetSettingsCharge_Charge_Voltage_Threshold((ushort)int.Parse(textBoxCharge_Voltage_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsCharge_Charge_Voltage_Recovery((ushort)int.Parse(textBoxCharge_Voltage_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsCharge_Charge_Completion_Time((byte)int.Parse(textBoxCharge_Completion_Time.Text));
            _batteryModule.GetMemory().SetSettingsCharge_Discharge_Voltage_Recovery((ushort)int.Parse(textBoxDischarge_Voltage_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsCharge_Discharge_Completion_Time((byte)int.Parse(textBoxDischarge_Completion_Time.Text));
            _batteryModule.GetMemory().SetSettingsCharge_Remaining_Capacity_Alarm_Percent((byte)int.Parse(textBoxRemaining_Capacity_Alarm_Percent.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV20_Threshold((ushort)int.Parse(textBoxCV20_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV20_Recovery((ushort)int.Parse(textBoxCV20_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV20_Time((byte)int.Parse(textBoxCV20_Time.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV40_Threshold((ushort)int.Parse(textBoxCV40_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV40_Recovery((ushort)int.Parse(textBoxCV40_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV40_Time((byte)int.Parse(textBoxCV40_Time.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV60_Threshold((ushort)int.Parse(textBoxCV60_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV60_Recovery((ushort)int.Parse(textBoxCV60_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV60_Time((byte)int.Parse(textBoxCV60_Time.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV80_Threshold((ushort)int.Parse(textBoxCV80_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV80_Recovery((ushort)int.Parse(textBoxCV80_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV80_Time((byte)int.Parse(textBoxCV80_Time.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV100_Threshold((ushort)int.Parse(textBoxCV100_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV100_Recovery((ushort)int.Parse(textBoxCV100_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsCharge_CV100_Time((byte)int.Parse(textBoxCV100_Time.Text));

            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcUpdateSettingsCharge);
        }

        private void buttonSetSettingsAlarm_Click(object sender, EventArgs e)
        {
            _batteryModule.GetMemory().SetSettingsAlarm_COVC_Threshold((ushort)int.Parse(textBoxCOVC_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_COVC_Recovery((ushort)int.Parse(textBoxCOVC_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_COVC_Time((byte)int.Parse(textBoxCOVC_Time.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_COVT_Threshold((ushort)int.Parse(textBoxCOVT_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_COVT_Recovery((ushort)int.Parse(textBoxCOVT_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_COVT_Time((byte)int.Parse(textBoxCOVT_Time.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_CUV_Threshold((ushort)int.Parse(textBoxCUV_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_CUV_Recovery((ushort)int.Parse(textBoxCUV_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_CUV_Time((byte)int.Parse(textBoxCUV_Time.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_COT_Threshold((sbyte)int.Parse(textBoxCOT_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_COT_Recovery((sbyte)int.Parse(textBoxCOT_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_COT_Time((byte)int.Parse(textBoxCOT_Time.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_CUT_Threshold((sbyte)int.Parse(textBoxCUT_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_CUT_Recovery((sbyte)int.Parse(textBoxCUT_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_CUT_Time((byte)int.Parse(textBoxCUT_Time.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_ROT_Threshold((sbyte)int.Parse(textBoxROT_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_ROT_Recovery((sbyte)int.Parse(textBoxROT_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_ROT_Time((byte)int.Parse(textBoxROT_Time.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_RUT_Threshold((sbyte)int.Parse(textBoxRUT_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_RUT_Recovery((sbyte)int.Parse(textBoxRUT_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_RUT_Time((byte)int.Parse(textBoxRUT_Time.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_UTC_Threshold((sbyte)int.Parse(textBoxUTC_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_UTC_Recovery((sbyte)int.Parse(textBoxUTC_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_UTC_Time((byte)int.Parse(textBoxUTC_Time.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_UTD_Threshold((sbyte)int.Parse(textBoxUTD_Threshold.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_UTD_Recovery((sbyte)int.Parse(textBoxUTD_Recovery.Text));
            _batteryModule.GetMemory().SetSettingsAlarm_UTD_Time((byte)int.Parse(textBoxUTD_Time.Text));

            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcUpdateSettingsAlarm);
        }

        private void buttonDefaultSettingsCalibration_Click(object sender, EventArgs e)
        {
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcDefaultCalibration);
            _batteryModule.ResetLastCommand();
        }

        private void buttonDefaultSettingsCharge_Click(object sender, EventArgs e)
        {
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcDefaultCharge);
            _batteryModule.ResetLastCommand();
        }

        private void buttonDefaultSettingsAlarm_Click(object sender, EventArgs e)
        {
            _batteryModule.SendCommandToBms(BatteryModule.UartCommand.UartCommandPcDefaultAlarm);
            _batteryModule.ResetLastCommand();
        }

        public void setPictureBoxRT_status(MemoryBms.RT_status status)
        {
            pictureBoxRT_status.Invoke((ThreadStart)delegate
            {
                pictureBoxRT_status.BackColor = (status == MemoryBms.RT_status.Transmit)
                    ? pictureBoxRed.BackColor
                    : pictureBoxGreen.BackColor;
            });
        }

        private void checkBoxWriteToExcel_CheckedChanged(object sender, EventArgs e)
        {
            _batteryModule.IsWriteToExcel = checkBoxWriteToExcel.Checked;
        }
    }
}
