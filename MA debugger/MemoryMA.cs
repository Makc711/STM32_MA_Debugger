using System.Runtime.InteropServices;

namespace MA_debugger
{
    class MemoryMa
    {
        public enum MaEvent : ushort
        {
            BufferEnablePos           = 0,
            BufferEnableMsk           = 1 << BufferEnablePos,        /*!< 0x0001 */
            BufferEnable              = BufferEnableMsk,             /*!< BUFF_PWR Normal */
            TransformerOutPos         = 1,
            TransformerOutMsk         = 1 << TransformerOutPos,      /*!< 0x0002 */
            TransformerOut            = TransformerOutMsk,           /*!< Transformer Out */
            BalancingInPos            = 2,
            BalancingInMsk            = 1 << BalancingInPos,         /*!< 0x0004 */
            BalancingIn               = BalancingInMsk,              /*!< Balancing_In Enable */
            BalancingOutPos           = 3,
            BalancingOutMsk           = 1 << BalancingOutPos,        /*!< 0x0008 */
            BalancingOut              = BalancingOutMsk,             /*!< Balancing_Out Enable */
            SafetyStatusCovPos        = 4,
            SafetyStatusCovMsk        = 1 << SafetyStatusCovPos,     /*!< 0x0010 */
            SafetyStatusCov           = SafetyStatusCovMsk,          /*!< Cell Overvoltage */
            SafetyStatusCuvPos        = 5,
            SafetyStatusCuvMsk        = 1 << SafetyStatusCuvPos,     /*!< 0x0020 */
            SafetyStatusCuv           = SafetyStatusCuvMsk,          /*!< Cell Undervoltage */
            SafetyStatusCotPos        = 6,
            SafetyStatusCotMsk        = 1 << SafetyStatusCotPos,     /*!< 0x0040 */
            SafetyStatusCot           = SafetyStatusCotMsk,          /*!< Cell Overtemperature */
            SafetyStatusCutPos        = 7,
            SafetyStatusCutMsk        = 1 << SafetyStatusCutPos,     /*!< 0x0080 */
            SafetyStatusCut           = SafetyStatusCutMsk,          /*!< Cell Undertemperature */
            SafetyStatusOttPos        = 8,
            SafetyStatusOttMsk        = 1 << SafetyStatusOttPos,     /*!< 0x0100 */
            SafetyStatusOtt           = SafetyStatusOttMsk,          /*!< Tansistor Overtemperature */
            SafetyStatusMaFailPos     = 9,
            SafetyStatusMaFailMsk     = 1 << SafetyStatusMaFailPos,  /*!< 0x0200 */
            SafetyStatusMaFail        = SafetyStatusMaFailMsk        /*!< MA circuit error */
        }

        public MaMeasurements Measurements { get; set; }
        public MaSettings Settings { get; set; }
        public byte[] SettingsBuffer { get; set; }
        private const int ChecksumConstant = 44111;

        public MemoryMa()
        {
            SettingsBuffer = new byte[Marshal.SizeOf(typeof(MaSettings))];
        }

        public byte CalculateSettingsChecksum()
        {
            ushort checksum = 0;
            foreach (var settingsByte in SettingsBuffer)
            {
                checksum += (ushort) (settingsByte * ChecksumConstant);
            }
            return (byte) checksum;
        }
    }

    struct MaMeasurements
    {
        public ushort U_cell;              /*!< Cell voltage, mV */
        public short  I_balance;           /*!< Balancing current, mA */
        public sbyte  TemperatureAnode;    /*!< Anode cell temperature (-), 'C */
        public sbyte  TemperatureCathode;  /*!< Cathode cell temperature (+), 'C */
        public sbyte  TemperatureVT1;      /*!< VT1 transistor temperature, 'C */
        public ushort MA_Event_Register;   /*!< MA Event Register */
    }

    struct MaSettings
    {
        public ushort COV_Threshold;   /*!< Cell Over Voltage Threshold, mV */
        public ushort COV_Recovery;    /*!< Cell Over Voltage Recovery, mV */
        public ushort COV_Time;        /*!< Cell Over Voltage Time, ms */
        public ushort CUV_Threshold;   /*!< Cell Under Voltage Threshold, mV */
        public ushort CUV_Recovery;    /*!< Cell Under Voltage Recovery, mV */
        public ushort CUV_Time;        /*!< Cell Under Voltage Time, ms */
        public sbyte  COT_Threshold;   /*!< Cell Over Temperature Threshold, 'C */
        public sbyte  COT_Recovery;    /*!< Cell Over Temperature Recovery, 'C */
        public ushort COT_Time;        /*!< Cell Over Temperature Time, ms */
        public sbyte  CUT_Threshold;   /*!< Cell Under Temperature Threshold, 'C */
        public sbyte  CUT_Recovery;    /*!< Cell Under Temperature Recovery, 'C */
        public ushort CUT_Time;        /*!< Cell Under Temperature Time, ms */
        public sbyte  OTT_Threshold;   /*!< Over Temperature Tansistor Threshold, 'C */
        public sbyte  OTT_Recovery;    /*!< Over Temperature Tansistor Recovery, 'C */
        public ushort OTT_Time;        /*!< Over Temperature Tansistor Time, ms */
    }
}
