using System;
using System.Runtime.InteropServices;

namespace MA_debugger
{
    public class MemoryBms
    {
        public enum CellStatus : ushort
        {
            Cell_Status_COVC_Pos           = (0),                               
            Cell_Status_COVC_Msk           = (1 << Cell_Status_COVC_Pos),       /*!< 0x0001 */
            Cell_Status_COVC               = Cell_Status_COVC_Msk,              /*!< Cell Overvoltage (Charger) */
            Cell_Status_COVT_Pos           = (1),                               
            Cell_Status_COVT_Msk           = (1 << Cell_Status_COVT_Pos),       /*!< 0x0002 */
            Cell_Status_COVT               = Cell_Status_COVT_Msk,              /*!< Cell Overvoltage (Transistor) */
            Cell_Status_CUV_Pos            = (2),                               
            Cell_Status_CUV_Msk            = (1 << Cell_Status_CUV_Pos),        /*!< 0x0004 */
            Cell_Status_CUV                = Cell_Status_CUV_Msk,               /*!< Cell Undervoltage */
            Cell_Status_COT_Pos            = (3),                                        
            Cell_Status_COT_Msk            = (1 << Cell_Status_COT_Pos),        /*!< 0x0008 */
            Cell_Status_COT                = Cell_Status_COT_Msk,               /*!< Cell Overtemperature */  // RESERVED !!!
            Cell_Status_CUT_Pos            = (4),                                         
            Cell_Status_CUT_Msk            = (1 << Cell_Status_CUT_Pos),        /*!< 0x0010 */
            Cell_Status_CUT                = Cell_Status_CUT_Msk,               /*!< Cell Undertemperature */  // RESERVED !!!
            Cell_Status_UTC_Pos            = (5),                                         
            Cell_Status_UTC_Msk            = (1 << Cell_Status_UTC_Pos),        /*!< 0x0020 */
            Cell_Status_UTC                = Cell_Status_UTC_Msk,               /*!< Cell Undertemperature (Charge) */  // RESERVED !!!
            Cell_Status_UTD_Pos            = (6),                                         
            Cell_Status_UTD_Msk            = (1 << Cell_Status_UTD_Pos),        /*!< 0x0040 */
            Cell_Status_UTD                = Cell_Status_UTD_Msk,               /*!< Cell Undertemperature (Discharge) */  // RESERVED !!!
            Cell_Status_CTMAX_Pos          = (7),                               
            Cell_Status_CTMAX_Msk          = (1 << Cell_Status_CTMAX_Pos),      /*!< 0x0080 */
            Cell_Status_CTMAX              = Cell_Status_CTMAX_Msk,             /*!< Maximum battery temperature */  // RESERVED !!!
            Cell_Status_CTMIN_Pos          = (8),                              
            Cell_Status_CTMIN_Msk          = (1 << Cell_Status_CTMIN_Pos),      /*!< 0x0100 */
            Cell_Status_CTMIN              = Cell_Status_CTMIN_Msk,             /*!< Minimum battery temperature */  // RESERVED !!!
            Cell_Status_CVMAX_Pos          = (9),                               
            Cell_Status_CVMAX_Msk          = (1 << Cell_Status_CVMAX_Pos),      /*!< 0x0200 */
            Cell_Status_CVMAX              = Cell_Status_CVMAX_Msk,             /*!< Maximum battery voltage */
            Cell_Status_CVMIN_Pos          = (10),                               
            Cell_Status_CVMIN_Msk          = (1 << Cell_Status_CVMIN_Pos),      /*!< 0x0400 */
            Cell_Status_CVMIN              = Cell_Status_CVMIN_Msk,             /*!< Minimum battery voltage */
            Cell_Status_Balancing_Pos      = (11),                               
            Cell_Status_Balancing_Msk      = (1 << Cell_Status_Balancing_Pos),  /*!< 0x0800 */
            Cell_Status_Balancing          = Cell_Status_Balancing_Msk         /*!< Balancing Enable */
        }

        public enum ChargeStatus : ushort
        {
            Charge_Status_CC_Pos = (0),
            Charge_Status_CC_Msk = (1 << Charge_Status_CC_Pos),       /*!< 0x0001 */
            Charge_Status_CC = Charge_Status_CC_Msk,                 /*!< Charge Complete */
            Charge_Status_DC_Pos = (1),
            Charge_Status_DC_Msk = (1 << Charge_Status_DC_Pos),       /*!< 0x0002 */
            Charge_Status_DC = Charge_Status_DC_Msk,                 /*!< Discharge Complete */
            Charge_Status_RCA_Pos = (2),
            Charge_Status_RCA_Msk = (1 << Charge_Status_RCA_Pos),      /*!< 0x0004 */
            Charge_Status_RCA = Charge_Status_RCA_Msk,                /*!< Remaining Capacity Alarm */
            Charge_Status_BAL_Pos = (3),
            Charge_Status_BAL_Msk = (1 << Charge_Status_BAL_Pos),      /*!< 0x0008 */
            Charge_Status_BAL = Charge_Status_BAL_Msk,                /*!< Balance */
            Charge_Status_SOC20_Pos = (4),
            Charge_Status_SOC20_Msk = (1 << Charge_Status_SOC20_Pos),    /*!< 0x0010 */
            Charge_Status_SOC20 = Charge_Status_SOC20_Msk,              /*!< State of Charge 20% */
            Charge_Status_SOC40_Pos = (5),
            Charge_Status_SOC40_Msk = (1 << Charge_Status_SOC40_Pos),    /*!< 0x0020 */
            Charge_Status_SOC40 = Charge_Status_SOC40_Msk,              /*!< State of Charge 40% */
            Charge_Status_SOC60_Pos = (6),
            Charge_Status_SOC60_Msk = (1 << Charge_Status_SOC60_Pos),    /*!< 0x0040 */
            Charge_Status_SOC60 = Charge_Status_SOC60_Msk,              /*!< State of Charge 60% */
            Charge_Status_SOC80_Pos = (7),
            Charge_Status_SOC80_Msk = (1 << Charge_Status_SOC80_Pos),    /*!< 0x0080 */
            Charge_Status_SOC80 = Charge_Status_SOC80_Msk,              /*!< State of Charge 80% */
            Charge_Status_SOC100_Pos = (8),
            Charge_Status_SOC100_Msk = (1 << Charge_Status_SOC100_Pos),   /*!< 0x0100 */
            Charge_Status_SOC100 = Charge_Status_SOC100_Msk             /*!< State of Charge 100% */
        }

        public enum SafetyStatus : ushort
        {
            Safety_Status_COVC_Pos = CellStatus.Cell_Status_COVC_Pos,
            Safety_Status_COVC_Msk = CellStatus.Cell_Status_COVC_Msk,   /*!< 0x0001 */
            Safety_Status_COVC = CellStatus.Cell_Status_COVC,           /*!< Cell Overvoltage (Charger) */
            Safety_Status_COVT_Pos = CellStatus.Cell_Status_COVT_Pos,
            Safety_Status_COVT_Msk = CellStatus.Cell_Status_COVT_Msk,   /*!< 0x0002 */
            Safety_Status_COVT = CellStatus.Cell_Status_COVT,           /*!< Cell Overvoltage (Transistor) */
            Safety_Status_CUV_Pos = CellStatus.Cell_Status_CUV_Pos,
            Safety_Status_CUV_Msk = CellStatus.Cell_Status_CUV_Msk,     /*!< 0x0004 */
            Safety_Status_CUV = CellStatus.Cell_Status_CUV,             /*!< Cell Undervoltage */
            Safety_Status_COT_Pos = CellStatus.Cell_Status_COT_Pos,
            Safety_Status_COT_Msk = CellStatus.Cell_Status_COT_Msk,     /*!< 0x0008 */
            Safety_Status_COT = CellStatus.Cell_Status_COT,             /*!< Cell Overtemperature */  // RESERVED !!!
            Safety_Status_CUT_Pos = CellStatus.Cell_Status_CUT_Pos,
            Safety_Status_CUT_Msk = CellStatus.Cell_Status_CUT_Msk,     /*!< 0x0010 */
            Safety_Status_CUT = CellStatus.Cell_Status_CUT,             /*!< Cell Undertemperature */  // RESERVED !!!
            Safety_Status_UTC_Pos = CellStatus.Cell_Status_UTC_Pos,
            Safety_Status_UTC_Msk = CellStatus.Cell_Status_UTC_Msk,     /*!< 0x0020 */
            Safety_Status_UTC = CellStatus.Cell_Status_UTC,             /*!< Undertemperature (Charge) */
            Safety_Status_UTD_Pos = CellStatus.Cell_Status_UTD_Pos,
            Safety_Status_UTD_Msk = CellStatus.Cell_Status_UTD_Msk,     /*!< 0x0040 */
            Safety_Status_UTD = CellStatus.Cell_Status_UTD,             /*!< Undertemperature (Discharge) */
            Safety_Status_ROT_Pos = (7),
            Safety_Status_ROT_Msk = (1 << Safety_Status_ROT_Pos),       /*!< 0x0080 */
            Safety_Status_ROT = Safety_Status_ROT_Msk,                  /*!< Radiator Overtemperature */
            Safety_Status_RUT_Pos = (8),
            Safety_Status_RUT_Msk = (1 << Safety_Status_RUT_Pos),       /*!< 0x0100 */
            Safety_Status_RUT = Safety_Status_RUT_Msk                   /*!< Radiator Undertemperature */
        }

        public enum BatteryStatus : ushort
        {
            Battery_Status_FC_Pos = (0),
            Battery_Status_FC_Msk = (1 << Battery_Status_FC_Pos),       /*!< 0x0001 */
            Battery_Status_FC = Battery_Status_FC_Msk,                 /*!< Fully Charged */
            Battery_Status_FD_Pos = (1),
            Battery_Status_FD_Msk = (1 << Battery_Status_FD_Pos),       /*!< 0x0002 */
            Battery_Status_FD = Battery_Status_FD_Msk,                /*!< Fully Discharged */
            Battery_Status_RCA_Pos = (2),
            Battery_Status_RCA_Msk = (1 << Battery_Status_RCA_Pos),      /*!< 0x0004 */
            Battery_Status_RCA = Battery_Status_RCA_Msk,                /*!< Remaining Capacity Alarm */
            Battery_Status_TDA_Pos = (3),
            Battery_Status_TDA_Msk = (1 << Battery_Status_TDA_Pos),      /*!< 0x0008 */
            Battery_Status_TDA = Battery_Status_TDA_Msk,                /*!< Terminate Discharge Alarm */
            Battery_Status_OTA_Pos = (4),
            Battery_Status_OTA_Msk = (1 << Battery_Status_OTA_Pos),      /*!< 0x0010 */
            Battery_Status_OTA = Battery_Status_OTA_Msk,                /*!< Over Temperature Alarm */
            Battery_Status_TCA_Pos = (5),
            Battery_Status_TCA_Msk = (1 << Battery_Status_TCA_Pos),      /*!< 0x0020 */
            Battery_Status_TCA = Battery_Status_TCA_Msk,                /*!< Terminate Charge Alarm */
            Battery_Status_OCA_Pos = (6),
            Battery_Status_OCA_Msk = (1 << Battery_Status_OCA_Pos),      /*!< 0x0040 */
            Battery_Status_OCA = Battery_Status_OCA_Msk,                /*!< Over Charged Alarm */
            Battery_Status_PS_Pos = (7),
            Battery_Status_PS_Msk = (1 << Battery_Status_PS_Pos),       /*!< 0x0080 */
            Battery_Status_PS = Battery_Status_PS_Msk                 /*!< Password Set */
        }

        public enum CellIndex
        {
            Cell_1,
            Cell_2,
            Cell_3,
            Cell_4,
            Temperature
        }

        public enum RT_status
        {
            Receive,
            Transmit
        }

        private SettingsCharge _settingsCharge;
        private SettingsAlarm _settingsAlarm;
        public StatusRegisters StatusRegisters { get; set; }

        public SettingsCharge SettingsCharge
        {
            get => _settingsCharge;
            set => _settingsCharge = value;
        }

        public SettingsAlarm SettingsAlarm
        {
            get => _settingsAlarm;
            set => _settingsAlarm = value;
        }

        //============== SetSettingsCharge ================
        public void SetSettingsCharge_Balance_Voltage_Threshold(ushort value)
        {
            _settingsCharge.Balance_Voltage_Threshold = value;
        }

        public void SetSettingsCharge_Balance_Voltage_Recovery(ushort value)
        {
            _settingsCharge.Balance_Voltage_Recovery = value;
        }

        public void SetSettingsCharge_Balance_Time(byte value)
        {
            _settingsCharge.Balance_Time = value;
        }

        public void SetSettingsCharge_Balance_Delta_Voltage(byte value)
        {
            _settingsCharge.Balance_Delta_Voltage = value;
        }

        public void SetSettingsCharge_Charge_Voltage_Threshold(ushort value)
        {
            _settingsCharge.Charge_Voltage_Threshold = value;
        }

        public void SetSettingsCharge_Charge_Voltage_Recovery(ushort value)
        {
            _settingsCharge.Charge_Voltage_Recovery = value;
        }

        public void SetSettingsCharge_Charge_Completion_Time(byte value)
        {
            _settingsCharge.Charge_Completion_Time = value;
        }

        public void SetSettingsCharge_Discharge_Voltage_Recovery(ushort value)
        {
            _settingsCharge.Discharge_Voltage_Recovery = value;
        }

        public void SetSettingsCharge_Discharge_Completion_Time(byte value)
        {
            _settingsCharge.Discharge_Completion_Time = value;
        }

        public void SetSettingsCharge_Remaining_Capacity_Alarm_Percent(byte value)
        {
            _settingsCharge.Remaining_Capacity_Alarm_Percent = value;
        }

        public void SetSettingsCharge_CV20_Threshold(ushort value)
        {
            _settingsCharge.CV20_Threshold = value;
        }

        public void SetSettingsCharge_CV20_Recovery(ushort value)
        {
            _settingsCharge.CV20_Recovery = value;
        }

        public void SetSettingsCharge_CV20_Time(byte value)
        {
            _settingsCharge.CV20_Time = value;
        }

        public void SetSettingsCharge_CV40_Threshold(ushort value)
        {
            _settingsCharge.CV40_Threshold = value;
        }

        public void SetSettingsCharge_CV40_Recovery(ushort value)
        {
            _settingsCharge.CV40_Recovery = value;
        }

        public void SetSettingsCharge_CV40_Time(byte value)
        {
            _settingsCharge.CV40_Time = value;
        }

        public void SetSettingsCharge_CV60_Threshold(ushort value)
        {
            _settingsCharge.CV60_Threshold = value;
        }

        public void SetSettingsCharge_CV60_Recovery(ushort value)
        {
            _settingsCharge.CV60_Recovery = value;
        }

        public void SetSettingsCharge_CV60_Time(byte value)
        {
            _settingsCharge.CV60_Time = value;
        }

        public void SetSettingsCharge_CV80_Threshold(ushort value)
        {
            _settingsCharge.CV80_Threshold = value;
        }

        public void SetSettingsCharge_CV80_Recovery(ushort value)
        {
            _settingsCharge.CV80_Recovery = value;
        }

        public void SetSettingsCharge_CV80_Time(byte value)
        {
            _settingsCharge.CV80_Time = value;
        }

        public void SetSettingsCharge_CV100_Threshold(ushort value)
        {
            _settingsCharge.CV100_Threshold = value;
        }

        public void SetSettingsCharge_CV100_Recovery(ushort value)
        {
            _settingsCharge.CV100_Recovery = value;
        }

        public void SetSettingsCharge_CV100_Time(byte value)
        {
            _settingsCharge.CV100_Time = value;
        }
        //=================================================

        //============== SetSettingsAlarm =================
        public void SetSettingsAlarm_COVC_Threshold(ushort value)
        {
            _settingsAlarm.COVC_Threshold = value;
        }

        public void SetSettingsAlarm_COVC_Recovery(ushort value)
        {
            _settingsAlarm.COVC_Recovery = value;
        }

        public void SetSettingsAlarm_COVC_Time(byte value)
        {
            _settingsAlarm.COVC_Time = value;
        }

        public void SetSettingsAlarm_COVT_Threshold(ushort value)
        {
            _settingsAlarm.COVT_Threshold = value;
        }

        public void SetSettingsAlarm_COVT_Recovery(ushort value)
        {
            _settingsAlarm.COVT_Recovery = value;
        }

        public void SetSettingsAlarm_COVT_Time(byte value)
        {
            _settingsAlarm.COVT_Time = value;
        }

        public void SetSettingsAlarm_CUV_Threshold(ushort value)
        {
            _settingsAlarm.CUV_Threshold = value;
        }

        public void SetSettingsAlarm_CUV_Recovery(ushort value)
        {
            _settingsAlarm.CUV_Recovery = value;
        }

        public void SetSettingsAlarm_CUV_Time(byte value)
        {
            _settingsAlarm.CUV_Time = value;
        }

        public void SetSettingsAlarm_COT_Threshold(sbyte value)
        {
            _settingsAlarm.COT_Threshold = value;
        }

        public void SetSettingsAlarm_COT_Recovery(sbyte value)
        {
            _settingsAlarm.COT_Recovery = value;
        }

        public void SetSettingsAlarm_COT_Time(byte value)
        {
            _settingsAlarm.COT_Time = value;
        }

        public void SetSettingsAlarm_CUT_Threshold(sbyte value)
        {
            _settingsAlarm.CUT_Threshold = value;
        }

        public void SetSettingsAlarm_CUT_Recovery(sbyte value)
        {
            _settingsAlarm.CUT_Recovery = value;
        }

        public void SetSettingsAlarm_CUT_Time(byte value)
        {
            _settingsAlarm.CUT_Time = value;
        }

        public void SetSettingsAlarm_ROT_Threshold(sbyte value)
        {
            _settingsAlarm.ROT_Threshold = value;
        }

        public void SetSettingsAlarm_ROT_Recovery(sbyte value)
        {
            _settingsAlarm.ROT_Recovery = value;
        }

        public void SetSettingsAlarm_ROT_Time(byte value)
        {
            _settingsAlarm.ROT_Time = value;
        }

        public void SetSettingsAlarm_RUT_Threshold(sbyte value)
        {
            _settingsAlarm.RUT_Threshold = value;
        }

        public void SetSettingsAlarm_RUT_Recovery(sbyte value)
        {
            _settingsAlarm.RUT_Recovery = value;
        }

        public void SetSettingsAlarm_RUT_Time(byte value)
        {
            _settingsAlarm.RUT_Time = value;
        }

        public void SetSettingsAlarm_UTC_Threshold(sbyte value)
        {
            _settingsAlarm.UTC_Threshold = value;
        }

        public void SetSettingsAlarm_UTC_Recovery(sbyte value)
        {
            _settingsAlarm.UTC_Recovery = value;
        }

        public void SetSettingsAlarm_UTC_Time(byte value)
        {
            _settingsAlarm.UTC_Time = value;
        }

        public void SetSettingsAlarm_UTD_Threshold(sbyte value)
        {
            _settingsAlarm.UTD_Threshold = value;
        }

        public void SetSettingsAlarm_UTD_Recovery(sbyte value)
        {
            _settingsAlarm.UTD_Recovery = value;
        }

        public void SetSettingsAlarm_UTD_Time(byte value)
        {
            _settingsAlarm.UTD_Time = value;
        }
        //=================================================
    }

    public struct StatusRegisters
    {
        public ushort firmwareVersion;
        public ushort batteryVoltage;
        public ushort cellVoltageMax;
        public ushort cellVoltageMin;
        public sbyte cellTempMax;
        public sbyte cellTempMin;
        public sbyte radiatorTemperature;
        public byte stateOfCharge;
        public ushort batteryStatus;
        public ushort safetyAlert;
        public ushort safetyStatus;
        public ushort chargeAlert;
        public ushort chargeStatus;
        public ushort ADCVoltage1;
        public ushort ADCVoltage2;
        public ushort ADCVoltage3;
        public ushort ADCVoltage4;
        public ushort cellVoltage1;
        public ushort cellVoltage2;
        public ushort cellVoltage3;
        public ushort cellVoltage4;
        public sbyte cellTemp1;
        public sbyte cellTemp2;
        public sbyte cellTemp3;
        public sbyte cellTemp4;
        public ushort cellStatus1;
        public ushort cellStatus2;
        public ushort cellStatus3;
        public ushort cellStatus4;
        public byte emptyByteToFullSize;   // RESERVED !!!
        public byte checksum;
        // Use the data placement rule to align!
    }

    public struct SettingsCharge
    {
        public ushort Balance_Voltage_Threshold; /*!< Voltage above which balancing resistors turn on, mV */
        public ushort Balance_Voltage_Recovery; /*!< Voltage below which balancing resistors turn off, mV */
        public byte Balance_Time; /*!< Minimum runtime balancing resistors, s */
        public byte Charge_Completion_Time; /*!< Storage time for setting the CHARGE_STATUS_CC bit (battery charge completed), s */
        public ushort Charge_Voltage_Threshold; /*!< The threshold of the event of reaching the minimum voltage of the element AB level of the full charge, mV */
        public ushort Charge_Voltage_Recovery; /*!< Threshold for resetting the event of reaching the minimum voltage of the element AB of the level of full charge, mV */
        public ushort Discharge_Voltage_Recovery; /*!< Threshold for resetting the event of reaching the minimum voltage of the element AB level of full discharge, mV */
        public byte Discharge_Completion_Time; /*!< Storage time for setting the bit CHARGE_STATUS_DC (battery fully discharged), s */
        public byte Remaining_Capacity_Alarm_Percent; /*!< The percentage of charge for the low charge signal CHARGE_STATUS_RCA, if zero, the function is disabled, % */
        public ushort CV20_Threshold; /*!< Cell Voltage 20% Threshold, mV */
        public ushort CV20_Recovery; /*!< Cell Voltage 20% Recovery, mV */
        public byte CV20_Time; /*!< Cell Voltage 20% Time, s */
        public byte CV40_Time; /*!< Cell Voltage 40% Time, s */
        public ushort CV40_Threshold; /*!< Cell Voltage 40% Threshold, mV */
        public ushort CV40_Recovery; /*!< Cell Voltage 40% Recovery, mV */
        public ushort CV60_Threshold; /*!< Cell Voltage 60% Threshold, mV */
        public ushort CV60_Recovery; /*!< Cell Voltage 60% Recovery, mV */
        public byte CV60_Time; /*!< Cell Voltage 60% Time, s */
        public byte CV80_Time; /*!< Cell Voltage 80% Time, s */
        public ushort CV80_Threshold; /*!< Cell Voltage 80% Threshold, mV */
        public ushort CV80_Recovery; /*!< Cell Voltage 80% Recovery, mV */
        public ushort CV100_Threshold; /*!< Cell Voltage 100% Threshold, mV */
        public ushort CV100_Recovery; /*!< Cell Voltage 100% Recovery,mV */
        public byte CV100_Time; /*!< Cell Voltage 100% Time, s */
        public byte Balance_Delta_Voltage; /* Voltage difference above which balancing resistors turn on, mV */
        // Use the data placement rule to align!
        // 40 bytes = 10  32-bits words.  It's - OK
        // !!! Full size (bytes) must be a multiple of 4 !!!
    }

    public struct SettingsAlarm
    {
        public ushort COVC_Threshold; /*!< Cell Over Voltage Charger Threshold, mV */
        public ushort COVC_Recovery; /*!< Cell Over Voltage Charger Recovery, mV */
        public byte COVC_Time; /*!< Cell Over Voltage Charger Time, s */
        public byte COVT_Time; /*!< Cell Over Voltage Transistor Time, s */
        public ushort COVT_Threshold; /*!< Cell Over Voltage Transistor Threshold, mV */
        public ushort COVT_Recovery; /*!< Cell Over Voltage Transistor Recovery, mV */
        public ushort CUV_Threshold; /*!< Cell Under Voltage Threshold, mV */
        public ushort CUV_Recovery; /*!< Cell Under Voltage Recovery, mV */
        public byte CUV_Time; /*!< Cell Under Voltage Time, s */
        public sbyte COT_Threshold; /*!< Cell Over Temperature Threshold, 'C */  // RESERVED !!!
        public sbyte COT_Recovery; /*!< Cell Over Temperature Recovery, 'C */  // RESERVED !!!
        public byte COT_Time; /*!< Cell Over Temperature Time, s */  // RESERVED !!!
        public sbyte CUT_Threshold; /*!< Cell Under Temperature Threshold, 'C */  // RESERVED !!!
        public sbyte CUT_Recovery; /*!< Cell Under Temperature Recovery, 'C */  // RESERVED !!!
        public byte CUT_Time; /*!< Cell Under Temperature Time, s */  // RESERVED !!!
        public sbyte ROT_Threshold; /*!< Radiator Over Temperature Threshold, 'C */
        public sbyte ROT_Recovery; /*!< Radiator Over Temperature Recovery, 'C */
        public byte ROT_Time; /*!< Radiator Over Temperature Time, s */
        public sbyte RUT_Threshold; /*!< Radiator Under Temperature Threshold, 'C */
        public sbyte RUT_Recovery; /*!< Radiator Under Temperature Recovery, 'C */
        public byte RUT_Time; /*!< Radiator Under Temperature Time, s */
        public sbyte UTC_Threshold; /*!< Under Temperature Charge Threshold, 'C */  // RESERVED !!!
        public sbyte UTC_Recovery; /*!< Under Temperature Charge Recovery, 'C */  // RESERVED !!!
        public byte UTC_Time; /*!< Under Temperature Charge Time, s */  // RESERVED !!!
        public sbyte UTD_Threshold; /*!< Under Temperature Discharge Threshold, 'C */  // RESERVED !!!
        public sbyte UTD_Recovery; /*!< Under Temperature Discharge Recovery, 'C */  // RESERVED !!!
        public byte UTD_Time; /*!< Under Temperature Discharge Time, s */  // RESERVED !!!
        public byte emptyByteToFullSize;  // RESERVED !!!
        public ushort emptyHalfWordToFullSize;  // RESERVED !!!
        // Use the data placement rule to align!
        // 36 byte = 9  32-bits words.  It's - OK
        // !!! Full size (bytes) must be a multiple of 4 !!!
    }
}
