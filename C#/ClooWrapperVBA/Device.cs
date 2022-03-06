using Cloo;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace ClooWrapperVBA
{
    /// <summary>
	/// ProgramDevice interface.
	/// </summary>
    [ComVisible(true)]
    [Guid("B571086A-A2C7-4886-A7B7-5D109DF62207")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDevice
    {
        /// <summary>
        /// Returns device name.
        /// </summary>
        [DispId(1), Description("Returns device name.")]
        string DeviceName { get; }

        /// <summary>
        /// Returns device type. (<see cref="ComputeDeviceTypes"/>)
        /// </summary>
        [DispId(2), Description("Returns device type.")]
        string DeviceType { get; }

        /// <summary>
        /// Returns vendor.
        /// </summary>
        [DispId(3), Description("Returns vendor.")]
        string DeviceVendor { get; }

        /// <summary>
        /// Returns availability state of device.
        /// </summary>
        [DispId(4), Description("Returns availability state of device.")]
        bool DeviceAvailable { get; }

        /// <summary>
        /// Returns device version.
        /// </summary>
        [DispId(5), Description("Returns device version.")]
        string DeviceVersion { get; }

        /// <summary>
        /// Returns device MaxComputeUnits.
        /// </summary>
        [DispId(6), Description("Returns device MaxComputeUnits.")]
        double MaxComputeUnits { get; }

        /// <summary>
        /// Returns device SingleCapabilites.
        /// </summary>
        [DispId(7), Description("Returns device SingleCapabilites.")]
        string SingleCapabilites { get; }

        /// <summary>
        /// Returns device AddressBits.
        /// </summary>
        [DispId(8), Description("Returns device AddressBits.")]
        int AddressBits { get; }

        /// <summary>
        /// Returns device CommandQueueFlags.
        /// </summary>
        [DispId(9), Description("Returns device CommandQueueFlags.")]
        string CommandQueueFlags { get; }

        /// <summary>
        /// Returns device CompilerAvailable.
        /// </summary>
        [DispId(10), Description("Returns device CompilerAvailable.")]
        bool CompilerAvailable { get; }

        /// <summary>
        /// Returns device DriverVersion.
        /// </summary>
        [DispId(11), Description("Returns device DriverVersion.")]
        string DriverVersion { get; }

        /// <summary>
        /// Returns device EndianLittle.
        /// </summary>
        [DispId(12), Description("Returns device EndianLittle.")]
        bool EndianLittle { get; }

        /// <summary>
        /// Returns device ErrorCorrectionSupport.
        /// </summary>
        [DispId(13), Description("Returns device ErrorCorrectionSupport.")]
        bool ErrorCorrectionSupport { get; }

        /// <summary>
        /// Returns device ExecutionCapabilities.
        /// </summary>
        [DispId(14), Description("Returns device ExecutionCapabilities.")]
        string ExecutionCapabilities { get; }

        /// <summary>
        /// Returns device Extensions.
        /// </summary>
        [DispId(15), Description("Returns device Extensions.")]
        string[] DeviceExtensions { get; }

        /// <summary>
        /// Returns device GlobalMemoryCacheLineSize.
        /// </summary>
        [DispId(16), Description("Returns device GlobalMemoryCacheLineSize.")]
        double GlobalMemoryCacheLineSize { get; }

        /// <summary>
        /// Returns device GlobalMemoryCacheSize.
        /// </summary>
        [DispId(17), Description("Returns device GlobalMemoryCacheSize.")]
        double GlobalMemoryCacheSize { get; }

        /// <summary>
        /// Returns device GlobalMemoryCacheType.
        /// </summary>
        [DispId(18), Description("Returns device GlobalMemoryCacheType.")]
        string GlobalMemoryCacheType { get; }

        /// <summary>
        /// Returns device GlobalMemorySize.
        /// </summary>
        [DispId(19), Description("Returns device GlobalMemorySize.")]
        double GlobalMemorySize { get; }

        /// <summary>
        /// Returns device HostUnifiedMemory.
        /// </summary>
        [DispId(20), Description("Returns device HostUnifiedMemory.")]
        bool HostUnifiedMemory { get; }

        /// <summary>
        /// Returns device Image2DMaxHeight.
        /// </summary>
        [DispId(21), Description("Returns device Image2DMaxHeight.")]
        double Image2DMaxHeight { get; }

        /// <summary>
        /// Returns device Image2DMaxWidth.
        /// </summary>
        [DispId(22), Description("Returns device Image2DMaxWidth.")]
        double Image2DMaxWidth { get; }

        /// <summary>
        /// Returns device Image3DMaxDepth.
        /// </summary>
        [DispId(23), Description("Returns device Image3DMaxDepth.")]
        double Image3DMaxDepth { get; }

        /// <summary>
        /// Returns device Image3DMaxHeight.
        /// </summary>
        [DispId(24), Description("Returns device Image3DMaxHeight.")]
        double Image3DMaxHeight { get; }

        /// <summary>
        /// Returns device Image3DMaxWidth.
        /// </summary>
        [DispId(25), Description("Returns device Image3DMaxWidth.")]
        double Image3DMaxWidth { get; }

        /// <summary>
        /// Returns device ImageSupport.
        /// </summary>
        [DispId(26), Description("Returns device ImageSupport.")]
        bool ImageSupport { get; }

        /// <summary>
        /// Returns device LocalMemorySize.
        /// </summary>
        [DispId(27), Description("Returns device LocalMemorySize.")]
        double LocalMemorySize { get; }

        /// <summary>
        /// Returns device LocalMemoryType.
        /// </summary>
        [DispId(28), Description("Returns device LocalMemoryType.")]
        string LocalMemoryType { get; }

        /// <summary>
        /// Returns device MaxClockFrequency.
        /// </summary>
        [DispId(29), Description("Returns device MaxClockFrequency.")]
        double MaxClockFrequency { get; }

        /// <summary>
        /// Returns device MaxConstantArguments.
        /// </summary>
        [DispId(30), Description("Returns device MaxConstantArguments.")]
        double MaxConstantArguments { get; }

        /// <summary>
        /// Returns device MaxConstantBufferSize.
        /// </summary>
        [DispId(31), Description("Returns device MaxConstantBufferSize.")]
        double MaxConstantBufferSize { get; }

        /// <summary>
        /// Returns device MaxMemoryAllocationSize.
        /// </summary>
        [DispId(32), Description("Returns device MaxMemoryAllocationSize.")]
        double MaxMemoryAllocationSize { get; }

        /// <summary>
        /// Returns device MaxParameterSize.
        /// </summary>
        [DispId(33), Description("Returns device MaxParameterSize.")]
        double MaxParameterSize { get; }

        /// <summary>
        /// Returns device MaxReadImageArguments.
        /// </summary>
        [DispId(34), Description("Returns device MaxReadImageArguments.")]
        double MaxReadImageArguments { get; }

        /// <summary>
        /// Returns device MaxSamplers.
        /// </summary>
        [DispId(35), Description("Returns device MaxSamplers.")]
        double MaxSamplers { get; }

        /// <summary>
        /// Returns device MaxWorkGroupSize.
        /// </summary>
        [DispId(36), Description("Returns device MaxWorkGroupSize.")]
        double MaxWorkGroupSize { get; }

        /// <summary>
        /// Returns device MaxWorkItemDimensions.
        /// </summary>
        [DispId(37), Description("Returns device MaxWorkItemDimensions.")]
        double MaxWorkItemDimensions { get; }

        /// <summary>
        /// Returns device MaxWorkItemSizes.
        /// </summary>
        [DispId(38), Description("Returns device MaxWorkItemSizes.")]
        double[] MaxWorkItemSizes { get; }

        /// <summary>
        /// Returns device MaxWriteImageArguments.
        /// </summary>
        [DispId(39), Description("Returns device MaxWriteImageArguments.")]
        double MaxWriteImageArguments { get; }

        /// <summary>
        /// Returns device MemoryBaseAddressAlignment.
        /// </summary>
        [DispId(40), Description("Returns device MemoryBaseAddressAlignment.")]
        double MemoryBaseAddressAlignment { get; }

        /// <summary>
        /// Returns device MinDataTypeAlignmentSize.
        /// </summary>
        [DispId(41), Description("Returns device MinDataTypeAlignmentSize.")]
        double MinDataTypeAlignmentSize { get; }

        /// <summary>
        /// Returns device NativeVectorWidthChar.
        /// </summary>
        [DispId(42), Description("Returns device NativeVectorWidthChar.")]
        double NativeVectorWidthChar { get; }

        /// <summary>
        /// Returns device NativeVectorWidthDouble.
        /// </summary>
        [DispId(43), Description("Returns device NativeVectorWidthDouble.")]
        double NativeVectorWidthDouble { get; }

        /// <summary>
        /// Returns device NativeVectorWidthFloat.
        /// </summary>
        [DispId(44), Description("Returns device NativeVectorWidthFloat.")]
        double NativeVectorWidthFloat { get; }

        /// <summary>
        /// Returns device NativeVectorWidthHalf.
        /// </summary>
        [DispId(45), Description("Returns device NativeVectorWidthHalf.")]
        double NativeVectorWidthHalf { get; }

        /// <summary>
        /// Returns device NativeVectorWidthInt.
        /// </summary>
        [DispId(46), Description("Returns device NativeVectorWidthInt.")]
        double NativeVectorWidthInt { get; }

        /// <summary>
        /// Returns device NativeVectorWidthLong.
        /// </summary>
        [DispId(47), Description("Returns device NativeVectorWidthLong.")]
        double NativeVectorWidthLong { get; }

        /// <summary>
        /// Returns device NativeVectorWidthShort.
        /// </summary>
        [DispId(48), Description("Returns device NativeVectorWidthShort.")]
        double NativeVectorWidthShort { get; }

        /// <summary>
        /// Returns device OpenCLCVersionString.
        /// </summary>
        [DispId(49), Description("Returns device OpenCLCVersionString.")]
        string OpenCLCVersionString { get; }

        /// <summary>
        /// Returns device PreferredVectorWidthChar.
        /// </summary>
        [DispId(50), Description("Returns device PreferredVectorWidthChar.")]
        double PreferredVectorWidthChar { get; }

        /// <summary>
        /// Returns device PreferredVectorWidthDouble.
        /// </summary>
        [DispId(51), Description("Returns device PreferredVectorWidthDouble.")]
        double PreferredVectorWidthDouble { get; }

        /// <summary>
        /// Returns device PreferredVectorWidthFloat.
        /// </summary>
        [DispId(52), Description("Returns device PreferredVectorWidthFloat.")]
        double PreferredVectorWidthFloat { get; }

        /// <summary>
        /// Returns device PreferredVectorWidthHalf.
        /// </summary>
        [DispId(53), Description("Returns device PreferredVectorWidthHalf.")]
        double PreferredVectorWidthHalf { get; }

        /// <summary>
        /// Returns device PreferredVectorWidthInt.
        /// </summary>
        [DispId(54), Description("Returns device PreferredVectorWidthInt.")]
        double PreferredVectorWidthInt { get; }

        /// <summary>
        /// Returns device PreferredVectorWidthLong.
        /// </summary>
        [DispId(55), Description("Returns device PreferredVectorWidthLong.")]
        double PreferredVectorWidthLong { get; }

        /// <summary>
        /// Returns device PreferredVectorWidthShort.
        /// </summary>
        [DispId(56), Description("Returns device PreferredVectorWidthShort.")]
        double PreferredVectorWidthShort { get; }

        /// <summary>
        /// Returns device Profile.
        /// </summary>
        [DispId(57), Description("Returns device Profile.")]
        string Profile { get; }

        /// <summary>
        /// Returns device ProfilingTimerResolution.
        /// </summary>
        [DispId(58), Description("Returns device ProfilingTimerResolution.")]
        double ProfilingTimerResolution { get; }

        /// <summary>
        /// Returns device VendorId.
        /// </summary>
        [DispId(59), Description("Returns device VendorId.")]
        double VendorId { get; }

        /// <summary>
        /// Error string.
        /// </summary>
        [DispId(60), Description("Error string.")]
        string ErrorString { get; set; }
    }

    /// <summary>
    /// Class Device (gets only configuration of device for defined platform).
    /// </summary>
    [Guid("F282B6B3-7F24-4E3A-AD14-FEEFF1E53513")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class Device : IDevice
    {
        private readonly int platformIndex = -1;
        private readonly int deviceIndex = -1;

        /// <summary>
        /// Error string.
        /// </summary>
        public string ErrorString { get; set; }

        /// <summary>
        /// Constructor of device.
        /// </summary>
        /// <param name="platformIndex">Platform index.</param>
        /// <param name="deviceIndex">Device index.</param>
        public Device(int platformIndex, int deviceIndex)
        {
            ErrorString = "";
            this.platformIndex = platformIndex;
            this.deviceIndex = deviceIndex;
        }

        /// <summary>
        /// Returns device name.
        /// </summary>
        public string DeviceName
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Name;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in DeviceName: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns device type. (<see cref="ComputeDeviceTypes"/>)
        /// </summary>
        public string DeviceType
        {
            get
            {
                try
                {
                    switch ((uint)ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Type)
                    {
                        case 1:
                            return "Default";

                        case 2:
                            return "CPU";

                        case 4:
                            return "GPU";

                        case 8:
                            return "Accelerator";

                        case 4294967295:
                            return "All";
                    }
                    return "";
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in DeviceType: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns vendor.
        /// </summary>
        public string DeviceVendor
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Vendor;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in DeviceVendor: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns availability state of device.
        /// </summary>
        public bool DeviceAvailable
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Available;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in DeviceAvailable: " + ex.Message;
                    return false;
                }
            }
        }

        /// <summary>
        /// Returns device version.
        /// </summary>
        public string DeviceVersion
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].VersionString;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in DeviceVersion: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns device MaxComputeUnits.
        /// </summary>
        public double MaxComputeUnits
        {
            get
            {
                try
                {
                    return (double)ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxComputeUnits;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxComputeUnits: " + ex.Message;
                    return -1;
                }
            }
        }

        /// <summary>
        /// Returns device SingleCapabilites.
        /// </summary>
        public string SingleCapabilites
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].SingleCapabilites.ToString();
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in SingleCapabilites: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns device AddressBits.
        /// </summary>
        public int AddressBits
        {
            get
            {
                try
                {
                    return (int)ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].AddressBits;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in AddressBits: " + ex.Message;
                    return -1;
                }
            }
        }

        /// <summary>
        /// Returns device CommandQueueFlags.
        /// </summary>
        public string CommandQueueFlags
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].CommandQueueFlags.ToString();
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in CommandQueueFlags: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns device CompilerAvailable.
        /// </summary>
        public bool CompilerAvailable
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].CompilerAvailable;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in CompilerAvailable: " + ex.Message;
                    return false;
                }
            }
        }

        /// <summary>
        /// Returns device DriverVersion.
        /// </summary>
        public string DriverVersion
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].DriverVersion;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in DriverVersion: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns device EndianLittle.
        /// </summary>
        public bool EndianLittle
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].EndianLittle;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in EndianLittle: " + ex.Message;
                    return false;
                }
            }
        }

        /// <summary>
        /// Returns device ErrorCorrectionSupport.
        /// </summary>
        public bool ErrorCorrectionSupport
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].ErrorCorrectionSupport;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in ErrorCorrectionSupport: " + ex.Message;
                    return false;
                }
            }
        }

        /// <summary>
        /// Returns device ExecutionCapabilities.
        /// </summary>
        public string ExecutionCapabilities
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].ExecutionCapabilities.ToString();
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in ExecutionCapabilities: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns device Extensions.
        /// </summary>
        public string[] DeviceExtensions
        {
            get
            {
                try
                {
                    string[] tmpStrings;
                    if (ComputePlatform.Platforms[platformIndex].Extensions.Count == 0)
                    {
                        tmpStrings = new string[1];
                    }
                    else
                    {
                        tmpStrings = new string[ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Extensions.Count];
                        ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Extensions.CopyTo(tmpStrings, 0);
                    }

                    return tmpStrings;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in DeviceExtensions: " + ex.Message;
                    return null;
                }
            }
        }

        /// <summary>
        /// Returns device GlobalMemoryCacheLineSize.
        /// </summary>
        public double GlobalMemoryCacheLineSize
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].GlobalMemoryCacheLineSize;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in GlobalMemoryCacheLineSize: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device GlobalMemoryCacheSize.
        /// </summary>
        public double GlobalMemoryCacheSize
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].GlobalMemoryCacheSize;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in GlobalMemoryCacheSize: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device GlobalMemoryCacheType.
        /// </summary>
        public string GlobalMemoryCacheType
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].GlobalMemoryCacheType.ToString();
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in GlobalMemoryCacheType: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns device GlobalMemorySize.
        /// </summary>
        public double GlobalMemorySize
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].GlobalMemorySize;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in GlobalMemorySize: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device HostUnifiedMemory.
        /// </summary>
        public bool HostUnifiedMemory
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].HostUnifiedMemory;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in HostUnifiedMemory: " + ex.Message;
                    return false;
                }
            }
        }

        /// <summary>
        /// Returns device Image2DMaxHeight.
        /// </summary>
        public double Image2DMaxHeight
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Image2DMaxHeight;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in Image2DMaxHeight: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device Image2DMaxWidth.
        /// </summary>
        public double Image2DMaxWidth
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Image2DMaxWidth;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in Image2DMaxWidth: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device Image3DMaxDepth.
        /// </summary>
        public double Image3DMaxDepth
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Image3DMaxDepth;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in Image3DMaxDepth: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device Image3DMaxHeight.
        /// </summary>
        public double Image3DMaxHeight
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Image3DMaxHeight;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in Image3DMaxHeight: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device Image3DMaxWidth.
        /// </summary>
        public double Image3DMaxWidth
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Image3DMaxWidth;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in Image3DMaxWidth: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device ImageSupport.
        /// </summary>
        public bool ImageSupport
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].ImageSupport;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in ImageSupport: " + ex.Message;
                    return false;
                }
            }
        }

        /// <summary>
        /// Returns device LocalMemorySize.
        /// </summary>
        public double LocalMemorySize
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].LocalMemorySize;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in LocalMemorySize: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device LocalMemoryType.
        /// </summary>
        public string LocalMemoryType
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].LocalMemoryType.ToString();
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in LocalMemoryType: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns device MaxClockFrequency.
        /// </summary>
        public double MaxClockFrequency
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxClockFrequency;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxClockFrequency: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MaxConstantArguments.
        /// </summary>
        public double MaxConstantArguments
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxConstantArguments;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxConstantArguments: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MaxConstantBufferSize.
        /// </summary>
        public double MaxConstantBufferSize
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxConstantBufferSize;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxConstantBufferSize: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MaxMemoryAllocationSize.
        /// </summary>
        public double MaxMemoryAllocationSize
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxMemoryAllocationSize;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxMemoryAllocationSize: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MaxParameterSize.
        /// </summary>
        public double MaxParameterSize
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxParameterSize;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxParameterSize: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MaxReadImageArguments.
        /// </summary>
        public double MaxReadImageArguments
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxReadImageArguments;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxReadImageArguments: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MaxSamplers.
        /// </summary>
        public double MaxSamplers
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxSamplers;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxSamplers: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MaxWorkGroupSize.
        /// </summary>
        public double MaxWorkGroupSize
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxWorkGroupSize;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxWorkGroupSize: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MaxWorkItemDimensions.
        /// </summary>
        public double MaxWorkItemDimensions
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxWorkItemDimensions;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxWorkItemDimensions: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MaxWorkItemSizes.
        /// </summary>
        public double[] MaxWorkItemSizes
        {
            get
            {
                try
                {
                    double[] maxWorkItemSizes = new double[ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxWorkItemSizes.Count];

                    for (int i = 0; i < maxWorkItemSizes.Length; i++)
                    {
                        maxWorkItemSizes[i] = ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxWorkItemSizes[i];
                    }

                    return maxWorkItemSizes;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxWorkItemSizes: " + ex.Message;
                    return null;
                }
            }
        }

        /// <summary>
        /// Returns device MaxWriteImageArguments.
        /// </summary>
        public double MaxWriteImageArguments
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MaxWriteImageArguments;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MaxWriteImageArguments: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MemoryBaseAddressAlignment.
        /// </summary>
        public double MemoryBaseAddressAlignment
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MemoryBaseAddressAlignment;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MemoryBaseAddressAlignment: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device MinDataTypeAlignmentSize.
        /// </summary>
        public double MinDataTypeAlignmentSize
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].MinDataTypeAlignmentSize;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in MinDataTypeAlignmentSize: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device NativeVectorWidthChar.
        /// </summary>
        public double NativeVectorWidthChar
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].NativeVectorWidthChar;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in NativeVectorWidthChar: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device NativeVectorWidthDouble.
        /// </summary>
        public double NativeVectorWidthDouble
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].NativeVectorWidthDouble;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in NativeVectorWidthDouble: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device NativeVectorWidthFloat.
        /// </summary>
        public double NativeVectorWidthFloat
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].NativeVectorWidthFloat;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in NativeVectorWidthFloat: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device NativeVectorWidthHalf.
        /// </summary>
        public double NativeVectorWidthHalf
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].NativeVectorWidthHalf;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in NativeVectorWidthHalf: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device NativeVectorWidthInt.
        /// </summary>
        public double NativeVectorWidthInt
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].NativeVectorWidthInt;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in NativeVectorWidthInt: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device NativeVectorWidthLong.
        /// </summary>
        public double NativeVectorWidthLong
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].NativeVectorWidthLong;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in NativeVectorWidthLong: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device NativeVectorWidthShort.
        /// </summary>
        public double NativeVectorWidthShort
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].NativeVectorWidthShort;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in NativeVectorWidthShort: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device OpenCLCVersionString.
        /// </summary>
        public string OpenCLCVersionString
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].OpenCLCVersionString;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in OpenCLCVersionString: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns device PreferredVectorWidthChar.
        /// </summary>
        public double PreferredVectorWidthChar
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].PreferredVectorWidthChar;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in PreferredVectorWidthChar: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device PreferredVectorWidthDouble.
        /// </summary>
        public double PreferredVectorWidthDouble
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].PreferredVectorWidthDouble;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in PreferredVectorWidthDouble: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device PreferredVectorWidthFloat.
        /// </summary>
        public double PreferredVectorWidthFloat
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].PreferredVectorWidthFloat;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in PreferredVectorWidthFloat: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device PreferredVectorWidthHalf.
        /// </summary>
        public double PreferredVectorWidthHalf
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].PreferredVectorWidthHalf;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in PreferredVectorWidthHalf: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device PreferredVectorWidthInt.
        /// </summary>
        public double PreferredVectorWidthInt
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].PreferredVectorWidthInt;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in PreferredVectorWidthInt: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device PreferredVectorWidthLong.
        /// </summary>
        public double PreferredVectorWidthLong
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].PreferredVectorWidthLong;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in PreferredVectorWidthLong: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device PreferredVectorWidthShort.
        /// </summary>
        public double PreferredVectorWidthShort
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].PreferredVectorWidthShort;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in PreferredVectorWidthShort: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device Profile.
        /// </summary>
        public string Profile
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Profile;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in Profile: " + ex.Message;
                    return "Error";
                }
            }
        }

        /// <summary>
        /// Returns device ProfilingTimerResolution.
        /// </summary>
        public double ProfilingTimerResolution
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].ProfilingTimerResolution;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in ProfilingTimerResolution: " + ex.Message;
                    return -1.0;
                }
            }
        }

        /// <summary>
        /// Returns device VendorId.
        /// </summary>
        public double VendorId
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].VendorId;
                }
                catch (Exception ex)
                {
                    ErrorString += "\r\nError in VendorId: " + ex.Message;
                    return -1.0;
                }
            }
        }
    }
}