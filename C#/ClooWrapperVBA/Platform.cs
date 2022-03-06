using Cloo;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ClooWrapperVBA
{
    [ComVisible(true)]
    [Guid("88ADB708-A83B-4A5A-8CB0-F3B708E32C1A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IPlatform
    {
        /// <summary>
        /// Returns a name of selected platform.
        /// </summary>
        [DispId(1), Description("Returns a name of selected platform.")]
        string PlatformName { get; }

        /// <summary>
        /// Returns a profile of selected platform.
        /// </summary>
        [DispId(2), Description("Returns a profile of selected platform.")]
        string PlatformProfile { get; }

        /// <summary>
        /// Returns a vendor of selected platform.
        /// </summary>
        [DispId(3), Description("Returns a vendor of selected platform.")]
        string PlatformVendor { get; }

        /// <summary>
        /// Returns an OpenCl version of selected platform.
        /// </summary>
        [DispId(4), Description("Returns an OpenCl version of selected platform.")]
        string PlatformVersion { get; }

        /// <summary>
        /// Returns extensions of selected platform.
        /// </summary>
        [DispId(5), Description("Returns extensions of selected platform.")]
        string[] PlatformExtensions { get; }

        /// <summary>
        /// Returns number of devices available for platformIndex.
        /// </summary>
        [DispId(6), Description("Returns number of devices available for platformIndex.")]
        int Devices { get; }

        /// <summary>
        /// Initialize device.
        /// </summary>
        /// <param name="deviceIndex">0-based device index.</param>
        /// <returns>True, if device was initialized successfully, otherwise false.</returns>
        [DispId(7), Description("Device initialization.")]
        bool SetDevice(int deviceIndex);

        /// <summary>
        /// Error message.
        /// </summary>
        [DispId(8), Description("Error message.")]
        string ErrorMessage { get; set; }

        /// <summary>
        /// Reference to device.
        /// </summary>
        [DispId(9), Description("Reference to device.")]
        Device Device { get; set; }
    }

    [Guid("7C0C3E18-6ECD-47C5-9BFD-92035099DA33")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class Platform : IPlatform
    {
        private readonly int platformIndex = -1;

        /// <summary>
        /// Constructor.
        /// </summary>
        public Platform()
        {
        }

        /// <summary>
        /// Constructor. Initializes platform.
        /// </summary>
        /// <param name="platformIndex">0-based platform index.</param>
        public Platform(int platformIndex)
        {
            this.platformIndex = platformIndex;
        }

        /// <summary>
        /// Returns a name of selected platform.
        /// </summary>
        public string PlatformName
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Name;
                }
                catch (Exception ex)
                {
                    ErrorMessage = ex.Message;
                    return "";
                }
            }
        }

        /// <summary>
        /// Returns a profile of selected platform.
        /// </summary>
        public string PlatformProfile
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Profile;
                }
                catch (Exception ex)
                {
                    ErrorMessage = ex.Message;
                    return "";
                }
            }
        }

        /// <summary>
        /// Returns a vendor of selected platform.
        /// </summary>
        public string PlatformVendor
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Vendor;
                }
                catch (Exception ex)
                {
                    ErrorMessage = ex.Message;
                    return "";
                }
            }
        }

        /// <summary>
        /// Returns an OpenCl version of selected platform.
        /// </summary>
        public string PlatformVersion
        {
            get
            {
                try
                {
                    return ComputePlatform.Platforms[platformIndex].Version;
                }
                catch (Exception ex)
                {
                    ErrorMessage = ex.Message;
                    return "";
                }
            }
        }

        /// <summary>
        /// Returns extensions of selected platform.
        /// </summary>
        public string[] PlatformExtensions
        {
            get
            {
                if (platformIndex < 0)
                {
                    MessageBox.Show(ErrorMessage);
                    return null;
                }
                else
                {
                    string[] tmpStrings;
                    if (ComputePlatform.Platforms[platformIndex].Extensions.Count == 0)
                    {
                        tmpStrings = new string[1];
                    }
                    else
                    {
                        tmpStrings = new string[ComputePlatform.Platforms[platformIndex].Extensions.Count];
                        ComputePlatform.Platforms[platformIndex].Extensions.CopyTo(tmpStrings, 0);
                    }

                    return tmpStrings;
                }
            }
        }

        /// <summary>
        /// Returns number of devices available for platformIndex.
        /// </summary>
        public int Devices
        {
            get
            {
                return ComputePlatform.Platforms[platformIndex].Devices.Count;
            }
        }

        /// <summary>
        /// Initialize device.
        /// </summary>
        /// <param name="deviceIndex">0-based device index.</param>
        /// <returns>True, if device was initialized successfully, otherwise false.</returns>
        public bool SetDevice(int deviceIndex)
        {
            if (platformIndex < ComputePlatform.Platforms.Count && deviceIndex < ComputePlatform.Platforms[platformIndex].Devices.Count)
            {
                Device = new Device(platformIndex, deviceIndex);
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Error message.
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// Device.
        /// </summary>
        public Device Device { get; set; } = null;
    }
}