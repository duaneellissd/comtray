using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace comtray
{

    /// <summary>
    ///  This "Scans" the Hardware, returning a list of serial ports sorted by name.
    /// </summary>
    public class ComportScanner
    {
        [Flags]
        private enum DiGetClassFlags : uint
        {
            DIGCF_DEFAULT = 0x00000001,  // only valid with DIGCF_DEVICEINTERFACE
            DIGCF_PRESENT = 0x00000002,
            DIGCF_ALLCLASSES = 0x00000004,
            DIGCF_PROFILE = 0x00000008,
            DIGCF_DEVICEINTERFACE = 0x00000010,
        }

        private const int utf16terminatorSize_bytes = 2;


        /// <summary>
        /// Device registry property codes
        /// </summary>
        private enum SPDRP : uint
        {
            /// <summary>
            /// DeviceDesc (R/W)
            /// </summary>
            SPDRP_DEVICEDESC = 0x00000000,

            /// <summary>
            /// HardwareID (R/W)
            /// </summary>
            SPDRP_HARDWAREID = 0x00000001,

            /// <summary>
            /// CompatibleIDs (R/W)
            /// </summary>
            SPDRP_COMPATIBLEIDS = 0x00000002,

            /// <summary>
            /// unused
            /// </summary>
            SPDRP_UNUSED0 = 0x00000003,

            /// <summary>
            /// Service (R/W)
            /// </summary>
            SPDRP_SERVICE = 0x00000004,

            /// <summary>
            /// unused
            /// </summary>
            SPDRP_UNUSED1 = 0x00000005,

            /// <summary>
            /// unused
            /// </summary>
            SPDRP_UNUSED2 = 0x00000006,

            /// <summary>
            /// Class (R--tied to ClassGUID)
            /// </summary>
            SPDRP_CLASS = 0x00000007,

            /// <summary>
            /// ClassGUID (R/W)
            /// </summary>
            SPDRP_CLASSGUID = 0x00000008,

            /// <summary>
            /// Driver (R/W)
            /// </summary>
            SPDRP_DRIVER = 0x00000009,

            /// <summary>
            /// ConfigFlags (R/W)
            /// </summary>
            SPDRP_CONFIGFLAGS = 0x0000000A,

            /// <summary>
            /// Mfg (R/W)
            /// </summary>
            SPDRP_MFG = 0x0000000B,

            /// <summary>
            /// FriendlyName (R/W)
            /// </summary>
            SPDRP_FRIENDLYNAME = 0x0000000C,

            /// <summary>
            /// LocationInformation (R/W)
            /// </summary>
            SPDRP_LOCATION_INFORMATION = 0x0000000D,

            /// <summary>
            /// PhysicalDeviceObjectName (R)
            /// </summary>
            SPDRP_PHYSICAL_DEVICE_OBJECT_NAME = 0x0000000E,

            /// <summary>
            /// Capabilities (R)
            /// </summary>
            SPDRP_CAPABILITIES = 0x0000000F,

            /// <summary>
            /// UiNumber (R)
            /// </summary>
            SPDRP_UI_NUMBER = 0x00000010,

            /// <summary>
            /// UpperFilters (R/W)
            /// </summary>
            SPDRP_UPPERFILTERS = 0x00000011,

            /// <summary>
            /// LowerFilters (R/W)
            /// </summary>
            SPDRP_LOWERFILTERS = 0x00000012,

            /// <summary>
            /// BusTypeGUID (R)
            /// </summary>
            SPDRP_BUSTYPEGUID = 0x00000013,

            /// <summary>
            /// LegacyBusType (R)
            /// </summary>
            SPDRP_LEGACYBUSTYPE = 0x00000014,

            /// <summary>
            /// BusNumber (R)
            /// </summary>
            SPDRP_BUSNUMBER = 0x00000015,

            /// <summary>
            /// Enumerator Name (R)
            /// </summary>
            SPDRP_ENUMERATOR_NAME = 0x00000016,

            /// <summary>
            /// Security (R/W, binary form)
            /// </summary>
            SPDRP_SECURITY = 0x00000017,

            /// <summary>
            /// Security (W, SDS form)
            /// </summary>
            SPDRP_SECURITY_SDS = 0x00000018,

            /// <summary>
            /// Device Type (R/W)
            /// </summary>
            SPDRP_DEVTYPE = 0x00000019,

            /// <summary>
            /// Device is exclusive-access (R/W)
            /// </summary>
            SPDRP_EXCLUSIVE = 0x0000001A,

            /// <summary>
            /// Device Characteristics (R/W)
            /// </summary>
            SPDRP_CHARACTERISTICS = 0x0000001B,

            /// <summary>
            /// Device Address (R)
            /// </summary>
            SPDRP_ADDRESS = 0x0000001C,

            /// <summary>
            /// UiNumberDescFormat (R/W)
            /// </summary>
            SPDRP_UI_NUMBER_DESC_FORMAT = 0X0000001D,

            /// <summary>
            /// Device Power Data (R)
            /// </summary>
            SPDRP_DEVICE_POWER_DATA = 0x0000001E,

            /// <summary>
            /// Removal Policy (R)
            /// </summary>
            SPDRP_REMOVAL_POLICY = 0x0000001F,

            /// <summary>
            /// Hardware Removal Policy (R)
            /// </summary>
            SPDRP_REMOVAL_POLICY_HW_DEFAULT = 0x00000020,

            /// <summary>
            /// Removal Policy Override (RW)
            /// </summary>
            SPDRP_REMOVAL_POLICY_OVERRIDE = 0x00000021,

            /// <summary>
            /// Device Install State (R)
            /// </summary>
            SPDRP_INSTALL_STATE = 0x00000022,

            /// <summary>
            /// Device Location Paths (R)
            /// </summary>
            SPDRP_LOCATION_PATHS = 0x00000023,
        }

        private const uint DICS_FLAG_GLOBAL = 0x00000001;
        private const uint DIREG_DEV = 0x00000001;
        private const uint KEY_QUERY_VALUE = 0x0001;

        /// <summary>
        /// The SP_DEVINFO_DATA structure defines a device instance that is a member of a device information set.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct SP_DEVINFO_DATA
        {
            public uint cbSize;
            public Guid ClassGuid;
            public uint DevInst;
            public UIntPtr Reserved;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DEVPROPKEY
        {
            public Guid fmtid;
            public uint pid;
        }

        [DllImport("setupapi.dll")]
        private static extern int SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiEnumDeviceInfo(IntPtr DeviceInfoSet, uint MemberIndex, ref SP_DEVINFO_DATA DeviceInterfaceData);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern IntPtr SetupDiGetClassDevs(ref Guid gClass, uint iEnumerator, uint hParent, DiGetClassFlags nFlags);

        [DllImport("Setupapi", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetupDiOpenDevRegKey(IntPtr hDeviceInfoSet, ref SP_DEVINFO_DATA deviceInfoData, uint scope,
            uint hwProfile, uint parameterRegistryValueKind, uint samDesired);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegQueryValueExW", SetLastError = true)]
        private static extern int RegQueryValueEx(IntPtr hKey, string lpValueName, int lpReserved, out uint lpType,
            byte[] lpData, ref uint lpcbData);

        [DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern int RegCloseKey(IntPtr hKey);

        [DllImport("kernel32.dll")]
        private static extern int GetLastError();

        private const int BUFFER_SIZE = 1024;

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiClassGuidsFromName(string ClassName,
            ref Guid ClassGuidArray1stItem, uint ClassGuidArraySize,
            out uint RequiredSize);

        [DllImport("setupapi.dll")]
        private static extern int SetupDiClassNameFromGuid(ref Guid ClassGuid,
            StringBuilder className, int ClassNameSize, ref int RequiredSize);

        /// <summary>
        /// The SetupDiGetDeviceRegistryProperty function retrieves the specified device property.
        /// This handle is typically returned by the SetupDiGetClassDevs or SetupDiGetClassDevsEx function.
        /// </summary>
        /// <param Name="DeviceInfoSet">Handle to the device information set that contains the interface and its underlying device.</param>
        /// <param Name="DeviceInfoData">Pointer to an SP_DEVINFO_DATA structure that defines the device instance.</param>
        /// <param Name="Property">Device property to be retrieved. SEE MSDN</param>
        /// <param Name="PropertyRegDataType">Pointer to a variable that receives the registry data Type. This parameter can be NULL.</param>
        /// <param Name="PropertyBuffer">Pointer to a buffer that receives the requested device property.</param>
        /// <param Name="PropertyBufferSize">Size of the buffer, in bytes.</param>
        /// <param Name="RequiredSize">Pointer to a variable that receives the required buffer size, in bytes. This parameter can be NULL.</param>
        /// <returns>If the function succeeds, the return value is nonzero.</returns>
        [DllImport("setupapi.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool SetupDiGetDeviceRegistryProperty(
            IntPtr DeviceInfoSet,
            ref SP_DEVINFO_DATA DeviceInfoData,
            SPDRP Property,
            out uint PropertyRegDataType,
            byte[] PropertyBuffer,
            uint PropertyBufferSize,
            out uint RequiredSize);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern bool SetupDiGetDevicePropertyW(
            IntPtr deviceInfoSet,
            [In] ref SP_DEVINFO_DATA DeviceInfoData,
            [In] ref DEVPROPKEY propertyKey,
            [Out] out uint propertyType,
            byte[] propertyBuffer,
            uint propertyBufferSize,
            out uint requiredSize,
            uint flags);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern int CM_Get_Parent(out UInt32 pdnDevInst, UInt32 dnDevInst, int ulFlags);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern int CM_Get_Device_IDW(UInt32 dnDevInst, byte[] buffer, uint buflen, uint flags);

        [DllImport("setupapi.dll", SetLastError = true)]
        private static extern int CM_Get_Device_ID_Size(out uint pulLen, UInt32 dnDevInst, int flags = 0);

        /// <summary>
        ///  This class holds USB details extracted from a Windows USB hardware ID.
        /// </summary>
        public class usb_info
        {
            public usb_info()
            {
                vid = 0; // 0 is invalid
                pid = 0; // 0 is valid
                mi = -1; // -1 not specified 
                serialnumber = "";
                hwid = "";
            }
            public int vid;
            public int pid;
            public int mi;
            public string serialnumber;
            public string hwid;
            // FUTURE: public string manufacturer;
            // FUTURE: public string product;
        }

        /// <summary>
        /// This class represents a UART Device of some sort.
        /// </summary>
        public class DeviceInfo
        {
            // name, ie: COM3
            public DeviceInfo()
            {
                name = "";
                description = "";
                usb_info = new usb_info();
                for_application_use = 0;
            }

            /// The quasi-devic name, ie: COM23
            public string name;
            // ie: Human friendly description
            public string description;
            // if usb, then vid is non-zero
            public usb_info usb_info;
            // for the client/application to use
            // The GUI uses this to track ADD/REMOVE
            public int for_application_use;
        }

        // used to sort by name, ie: COM3 vrs COM33 
        static public int CompareDeviceInfo(DeviceInfo A, DeviceInfo B)
        {
            int r;
            r = A.name.Length - B.name.Length;
            if( r != 0 ) { 
                /* Simple STRING CMP will put  COM200 before COM2
                 * We don't want that so first compare is by length.
                 * Since COM2 is shorter then COM200, COM2 wins!
                 */
                if( r < 0)
                {
                    return -1;
                } else
                {
                    return +1;
                }

            }
            /* Since the LENGTH is the same, we can sort by name 
             * Exampel:COM201 is before COM202 */
            r = A.name.CompareTo(B.name);
            if( r != 0)
            {
                return r;
            }
            r = A.usb_info.hwid.CompareTo(B.usb_info.hwid);
            return r;
        }

        /// <summary>
        /// This is the public Scanner Function that returns the list of comports present.
        /// </summary>
        /// <returns></returns>
        public static List<DeviceInfo> GetComportList()
        {
            // Per PYSERIAL - There are two different GUIdes to scan.
            string[] gnames = { "Modems", "Ports" };
            List<DeviceInfo> devices = new List<DeviceInfo>();
            uint iMemberIndex;
            IntPtr hDeviceInfoSet;
            hDeviceInfoSet = new IntPtr(0);

            // For each named Guid.
            foreach (string gn in gnames)
            {
                // Get that guid from windows
                Guid[] these_guids = GetClassGUIDs(gn);

                // for each of those guids..
                foreach (Guid g in these_guids)
                {
                    Guid gr = g;
                    /* start the scan process for this guid */
                    hDeviceInfoSet = SetupDiGetClassDevs(ref gr, 0, 0, DiGetClassFlags.DIGCF_PRESENT);
                    if (hDeviceInfoSet == IntPtr.Zero)
                        continue; /* try next guid */

                    // index through the found ports.
                    iMemberIndex = 0;
                    while (true)
                    {
                        SP_DEVINFO_DATA deviceInfoData = new SP_DEVINFO_DATA();
                        deviceInfoData.cbSize = (uint)Marshal.SizeOf(typeof(SP_DEVINFO_DATA));

                        /* Sort of "open" the device so we can ask questions about the device */
                        bool success = SetupDiEnumDeviceInfo(hDeviceInfoSet, iMemberIndex, ref deviceInfoData);
                        if (!success)
                        {
                            // No more devices in the device information set
                            break;
                        }

                        // create the device */
                        DeviceInfo deviceInfo = new DeviceInfo();

                        // get its name
                        deviceInfo.name = GetDeviceName(hDeviceInfoSet, deviceInfoData);
                        // and the Friendly description
                        // Example: STMicroelectronics STLink Virtual COM Port
                        deviceInfo.description = GetDeviceDescription(hDeviceInfoSet, deviceInfoData);

                        // USB details are more complicated to get 
                        GetUsbDetails(hDeviceInfoSet, deviceInfoData.DevInst, ref deviceInfo);

                        // PORTS = In windows terms = Printer + Serials.
                        // we don't want printers.
                        if( ! deviceInfo.name.StartsWith("LPT"))
                        {
                            devices.Add(deviceInfo);
                        }

                        // next device index
                        iMemberIndex++;
                    }
                }
                // close the scanning operation
                SetupDiDestroyDeviceInfoList(hDeviceInfoSet);
            }

            /* return the ports in sorted order.*/
            devices.Sort(CompareDeviceInfo);
            return devices;
        }

        /// <summary>
        ///  Given a USB Hardware ID, break it down and parse out details
        /// </summary>
        /// <param name="hwid"></param>
        /// <returns></returns>
        private static usb_info parse_usb(string hwid)
        {
            int idx;
            int tmp;
            usb_info info = new usb_info();

            // This Example comes from an STM32F103 Nucleo board.
            //
            // Example:  hwid = "USB\\VID_0483&PID_374B&MI_02\\7&2F22E1E4&0&0002\0"
            //  This is Vendorid 0x0483, Product id 0x374b, Interface 2
            //
            // The "crap" after the MI_02(interface) is ephemrail information that
            // is of no value to our needs.
            //
            // SOMETIMES .. we have to get a parent of this USB device.
            //
            // In this case, we get: hwid = "USB\\VID_0483&PID_374B\\0673FF494951785087021535\0"
            // Which is the same VID/PID - but the SERIAL NUMBER is present! YEA!
            //
            // And because it is ARM CMSIS, it contains a pseudo thumb drive.
            // and thumb drives require by specification a HEX string for a serial number.
            //
            // Other times it is just some string... with any combination of letters and stuff.
            //
            // We give up if we have gone up the parent/child stack 4 times.
            info.hwid = hwid;

            idx = hwid.IndexOf("VID_");
            if (idx < 0)
                return info;
            info.vid = Convert.ToInt32(hwid.Substring(idx + 4, 4), 16);

            idx = hwid.IndexOf("PID_", idx);
            if (idx < 0)
            {
                info.pid = 0;
                return info;
            }
            info.pid = Convert.ToInt32(hwid.Substring(idx + 4, 4), 16);
            idx = idx + 8;

            // the interface might not be present
            tmp = hwid.IndexOf("&MI_", idx);
            if (tmp > 0)
            {
                info.mi = Convert.ToInt32(hwid.Substring(tmp + 4, 2), 10);
                idx = tmp + 6;
            }

            // USB stuff works this way
            if (!hwid.StartsWith("USB\\"))
            {
                // if it starts with FTDI .. we'll do something else
                return info;
            }
            if (hwid[idx] != '\\')
            {
                return info;
            }

            while (true)
            {
                char c;
                idx++;
                c = hwid[idx];
                if (Convert.ToInt32(c) == 0)
                {
                    break;
                }
                if (char.IsLetterOrDigit(c))
                {
                    info.serialnumber += c;
                    continue;
                }
                if (c == '_')
                {
                    /* WINC Cell Modem, AVENT eval board
                     * hwid = "USB\\VID_1435&PID_3182\\WNC_ADB
                     * VID =  Wistron NeWeb Corp
                     */
                    info.serialnumber += c;
                    continue;
                }
                /* garbage */
                info.serialnumber = "";
                break;
            }


            return info;
        }

        // The device hw string is the windows string that describes the device.
        private static string GetDeviceHwString(uint deviceid)
        {
            uint ilen;
            int r;
            string s;

            CM_Get_Device_ID_Size(out ilen, deviceid, 0);
            ilen *= 2;
            ilen += 2;
            byte[] ptrBuf = new byte[ilen];
            r = CM_Get_Device_IDW(deviceid, ptrBuf, ilen, 0);
            s = Encoding.Unicode.GetString(ptrBuf);
            return s;
        }


        private static void GetUsbDetails(IntPtr hDeviceInfoSet, uint deviceid, ref DeviceInfo pDevInfo)
        {
            uint parentid;
            usb_info parent_info;
            string s;


            s = GetDeviceHwString(deviceid);
            /* not a usb device? */
            if (!(s.StartsWith("USB") || s.StartsWith("FTDIBUS")))
            {
                return;
            }
            pDevInfo.usb_info = parse_usb(s);

            // Did that give us a Serial Number? If so great we are done.
            if (pDevInfo.usb_info.serialnumber != "")
                return;

            // Go up the parent/child stack until we find a serial number
            for (int lvl = 0; lvl < 4; lvl++)
            {
                int r;
                r = CM_Get_Parent(out parentid, deviceid, 0);
                if (r != 0)
                {
                    // no more.. we are done.
                    break;
                }

                deviceid = parentid;
                /* get parent hw string */
                s = GetDeviceHwString(deviceid);

                // parse out USB stuff
                parent_info = parse_usb(s);

                // if not a match .. we have left our device.
                // Example: We have left a device and found the USB hub.
                // Verses:
                // Example: We have a DUAL comport device, we found 1 comport.
                //          We want the PARENT device of the DUAL comport device.
                //
                // The VID/PID should match.
                if (parent_info.vid != pDevInfo.usb_info.vid)
                    break;
                if (parent_info.pid != pDevInfo.usb_info.pid)
                    break;

                /* ok we found it, did we get a serial number? */
                pDevInfo.usb_info.serialnumber = parent_info.serialnumber;

                // And in somecases if we still don't have an interface.. 
                // but the parent has an interface we take that interface.
                if ((pDevInfo.usb_info.mi < 0) && (parent_info.mi >= 0))
                {
                    pDevInfo.usb_info.mi = parent_info.mi;
                }

                // keep climbing the stack until we find a serial number
                if (pDevInfo.usb_info.serialnumber != "")
                {
                    break;
                }
            }
        }

        /// Device name = COM243 or simular.
        private static string GetDeviceName(IntPtr pDevInfoSet, SP_DEVINFO_DATA deviceInfoData)
        {
            IntPtr hDeviceRegistryKey = SetupDiOpenDevRegKey(pDevInfoSet, ref deviceInfoData, DICS_FLAG_GLOBAL, 0, DIREG_DEV, KEY_QUERY_VALUE);

            if (hDeviceRegistryKey == IntPtr.Zero)
                return string.Empty; //Failed to open a registry key for device-specific configuration information

            byte[] ptrBuf = new byte[BUFFER_SIZE];
            uint length = (uint)ptrBuf.Length;
            try
            {
                uint lpRegKeyType;
                int result = RegQueryValueEx(hDeviceRegistryKey, "PortName", 0, out lpRegKeyType, ptrBuf, ref length);

                if (result == 0)
                    return Encoding.Unicode.GetString(ptrBuf, 0, (int)length - utf16terminatorSize_bytes);
            }
            finally
            {
                RegCloseKey(hDeviceRegistryKey);
            }

            return string.Empty; //Can not read registry value PortName for device
        }

        // The Human long description
        private static string GetDeviceDescription(IntPtr hDeviceInfoSet, SP_DEVINFO_DATA deviceInfoData)
        {
            byte[] ptrBuf = new byte[BUFFER_SIZE];
            uint propRegDataType;
            uint RequiredSize;
            bool success = SetupDiGetDeviceRegistryProperty(hDeviceInfoSet, ref deviceInfoData, SPDRP.SPDRP_DEVICEDESC,
                out propRegDataType, ptrBuf, BUFFER_SIZE, out RequiredSize);

            if (success)
                return Encoding.Unicode.GetString(ptrBuf, 0, (int)RequiredSize - utf16terminatorSize_bytes);

            return string.Empty; //Can not read registry value PortName for device
        }

        // Given a CLASS name get its GUID.
        private static Guid[] GetClassGUIDs(string className)
        {
            uint requiredSize;
            Guid[] guidArray = new Guid[1];

            bool status = SetupDiClassGuidsFromName(className, ref guidArray[0], 1, out requiredSize);
            if (status)
            {
                if (1 < requiredSize)
                {
                    guidArray = new Guid[requiredSize];
                    SetupDiClassGuidsFromName(className, ref guidArray[0], requiredSize, out requiredSize);
                }
            }

            return guidArray;
        }
    }
}
