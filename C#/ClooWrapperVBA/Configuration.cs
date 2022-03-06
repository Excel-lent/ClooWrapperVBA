using Cloo;
using System;
using System.Runtime.InteropServices;

namespace ClooWrapperVBA
{
    [ComVisible(true)]
    [Guid("EFE81401-D294-4377-832B-1E00AB0AB978")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Configuration
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        public Configuration()
        {
        }

        /// <summary>
        /// Platform.
        /// </summary>
        public Platform Platform = null;

        #region Platform-dependent properties and methods

        /// <summary>
        /// Returns a number of available platforms.
        /// </summary>
        public int Platforms
        {
            get
            {
                return ComputePlatform.Platforms.Count;
            }
        }

        /// <summary>
        /// Initialize platform.
        /// </summary>
        /// <param name="platformIndex">0-based platform index.</param>
        /// <returns>True, if platform was initialized successfully, otherwise false.</returns>
        public bool SetPlatform(int platformIndex)
        {
            if (platformIndex < ComputePlatform.Platforms.Count)
            {
                Platform = new Platform(platformIndex);
                return true;
            }
            else
            {
                return false;
            }
        }

        #endregion Platform-dependent properties and methods
    }
}