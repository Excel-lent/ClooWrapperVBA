using Cloo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;

namespace ClooWrapperVBA
{
    /// <summary>
    /// ProgramDevice interface.
    /// </summary>
    [ComVisible(true)]
    [Guid("2BF7DA6B-DDB3-42A5-BD65-92EE93ABB473")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IProgramDevice
    {
        /// <summary>
        /// Creates kernel for method.
        /// </summary>
        /// <param name="method">Method.</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        [DispId(1)]
        bool CreateKernel(string method);

        /// <summary>
        /// Initializes device of selected <paramref name="platformIndex"/> and <paramref name="deviceIndex"/>.
        /// Loads sources and compiles them.
        /// </summary>
        /// <param name="sourceCode">Sources as plain text.</param>
        /// <param name="options">Compilation options.</param>
        /// <param name="platformIndex">Platform index (<see cref="Configuration"/>).</param>
        /// <param name="deviceIndex">Device index (<see cref="Configuration"/>).</param>
        /// <param name="deviceTypeIndex">Device index inside one device type that corresponds to device index (for example, more than 1 "GPU" can be 
        /// installed on the current platform).</param>
        /// <param name="buildLogs">Build logs as single string.</param>
        /// <returns>True, if the sources were compiled successfully, false otherwise.</returns>
        [DispId(2)]
        bool Build(string sourceCode, string options, int platformIndex, int deviceIndex, int deviceTypeIndex, out string buildLogs);

        #region SetArguments

        /// <summary>
        /// Writes an array of type "Long" to the device.
        /// Be careful: The sequence of "SetMemoryArgument" must correspond to the sequence of argument in the method!
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="values">Array of "Long".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        [DispId(3)]
        bool SetMemoryArgument_Long(int argument_index, ref int[] values);

        /// <summary>
        /// Writes an array of type "Single" to the device.
        /// Be careful: The sequence of "SetMemoryArgument" must correspond to the sequence of argument in the method!
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="values">Array of "Single".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        [DispId(4)]
        bool SetMemoryArgument_Single(int argument_index, ref float[] values);

        /// <summary>
        /// Writes an array of type "Double" to the device.
        /// Be careful: The sequence of "SetMemoryArgument" must correspond to the sequence of argument in the method!
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="values">Array of "Double".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        [DispId(5)]
        bool SetMemoryArgument_Double(int argument_index, ref double[] values);

        /// <summary>
        /// Sets "Long" argument to the kernel.
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="value_long">Argument value as "Long".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        [DispId(6)]
        bool SetValueArgument_Long(int argument_index, int value_long);

        /// <summary>
        /// Sets "Single" argument to the kernel.
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="value_single">Argument value as "Single".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        [DispId(7)]
        bool SetValueArgument_Single(int argument_index, float value_single);

        /// <summary>
        /// Sets "Double" argument to the kernel.
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="value_double">Argument value as "Double".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        [DispId(8)]
        bool SetValueArgument_Double(int argument_index, double value_double);

        #endregion SetArguments

        #region Execution

        /// <summary>
        /// Synchronous execution.
        /// </summary>
        /// <param name="globalWorkOffset">Array of global work offset, or "null".</param>
        /// <param name="globalWorkSize">Array of global work size, or "null".</param>
        /// <param name="localWorkSize">Array of local work size, or "null".</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        [DispId(9)]
        bool ExecuteSync(ref int[] globalWorkOffset, ref int[] globalWorkSize, ref int[] localWorkSize);

        /// <summary>
        /// Execution in background.
        /// </summary>
        /// <param name="globalWorkOffset">Array of global work offset, or "null".</param>
        /// <param name="globalWorkSize">Array of global work size, or "null".</param>
        /// <param name="localWorkSize">Array of local work size, or "null".</param>
        /// <param name="threadPriority">Thread priority as integer (0 - "Lowest", 1 - "BelowNormal", 2 - "Normal", 3 - "AboveNormal", 4 - "Highest").</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        [DispId(10)]
        bool ExecuteBackground(ref int[] globalWorkOffset, ref int[] globalWorkSize, ref int[] localWorkSize, int threadPriority);

        /// <summary>
        /// For asynchronous call from VBA we need an address of callback function because VBA can use events only
        /// from form/classes.
        /// </summary>
        /// <param name="globalWorkOffset">Array of global work offset, or "null".</param>
        /// <param name="globalWorkSize">Array of global work size, or "null".</param>
        /// <param name="localWorkSize">Array of local work size, or "null".</param>
        /// <param name="threadPriority">Thread priority as integer (0 - "Lowest", 1 - "BelowNormal", 2 - "Normal", 3 - "AboveNormal", 4 - "Highest").</param>
        /// <param name="callback">Callback to the VBA function.</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        [DispId(11)]
        bool ExecuteAsync(ref int[] globalWorkOffset, ref int[] globalWorkSize, ref int[] localWorkSize, int threadPriority,
            [MarshalAs(UnmanagedType.FunctionPtr)] ref Action callback);

        /// <summary>
        /// True, if execution is completed, false otherwise.
        /// </summary>
        [DispId(12)]
        bool ExecutionCompleted { get; set; }

        #endregion Execution

        #region GetArguments

        /// <summary>
        /// Reads an array of type "Long" from the device.
        /// </summary>
        /// <param name="varIndex">0-based number of argument in argument list.</param>
        /// <param name="values">Array of "Long".</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        [DispId(13)]
        bool GetMemoryArgument_Long(int varIndex, ref int[] values);

        /// <summary>
        /// Reads an array of type "Single" from the device.
        /// </summary>
        /// <param name="varIndex">0-based number of argument in argument list.</param>
        /// <param name="values">Array of "Single".</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        [DispId(14)]
        bool GetMemoryArgument_Single(int varIndex, ref float[] values);

        /// <summary>
        /// Reads an array of type "Double" from the device.
        /// </summary>
        /// <param name="varIndex">0-based number of argument in argument list.</param>
        /// <param name="values">Array of "Double".</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        [DispId(15)]
        bool GetMemoryArgument_Double(int varIndex, ref double[] values);

        #endregion GetArguments

        /// <summary>
        /// Device type of used device ("GPU" / "CPU").
        /// </summary>
        [DispId(16)]
        string DeviceType { get; set; }

        /// <summary>
        /// Error string.
        /// </summary>
        [DispId(17)]
        string ErrorString { get; set; }
    }

    [ComVisible(true)]
    [Guid("56C41646-10CB-4188-979D-23F70E0FFDF5")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ProgramDevice : IProgramDevice
    {
        public ComputeProgram Prog;
        public ComputeContext ComputeContext;
        public ComputeCommandQueue ComputeCommandQueue = null;
        private ComputeKernel kernel;
        private Dictionary<int, ComputeMemory> variablePointers;
        private Action callBack;
        private long[] _globalWorkOffset = null;
        private long[] _globalWorkSize = null;
        private long[] _localWorkSize = null;

        /// <summary>
        /// Creates kernel for method.
        /// </summary>
        /// <param name="method">Method.</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        public bool CreateKernel(string method)
        {
            try
            {
                kernel = Prog.CreateKernel(method);
                variablePointers = new Dictionary<int, ComputeMemory>();
                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in CreateKernel: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// Initializes device of selected <paramref name="platformIndex"/> and <paramref name="deviceIndex"/>.
        /// Loads sources and compiles them.
        /// </summary>
        /// <param name="sourceCode">Sources as plain text.</param>
        /// <param name="options">Compilation options.</param>
        /// <param name="platformIndex">Platform index (<see cref="Configuration"/>).</param>
        /// <param name="deviceIndex">Device index (<see cref="Configuration"/>).</param>
        /// <param name="deviceTypeIndex">Device index inside one device type that corresponds to device index (for example, more than 1 "GPU" can be 
        /// installed on the current platform).</param>
        /// <param name="buildLogs">Build logs as single string.</param>
        /// <returns>True, if the sources were compiled successfully, false otherwise.</returns>
        public bool Build(string sourceCode, string options, int platformIndex, int deviceIndex, int deviceTypeIndex, out string buildLogs)
        {
            buildLogs = "";

            Device device = new Device(platformIndex, deviceIndex);
            DeviceType = device.DeviceType;

            if (!device.CompilerAvailable)
            {
                buildLogs = "Compiler is not available for selected device.";
                return false;
            }

            try
            {
                ComputeContext = new ComputeContext(ComputePlatform.Platforms[platformIndex].Devices[deviceIndex].Type,
                    new ComputeContextPropertyList(ComputePlatform.Platforms[platformIndex]), null, IntPtr.Zero);
            }
            catch (Exception e)
            {
                buildLogs = "Error during creation of ComputeContext:\r\n";
                buildLogs += e.Message;
                return false;
            }

            try
            {
                ComputeCommandQueue = new ComputeCommandQueue(ComputeContext, ComputeContext.Devices[deviceTypeIndex], ComputeCommandQueueFlags.None);
            }
            catch (Exception e)
            {
                buildLogs = "Error during creation of ComputeCommandQueue:\r\n";
                buildLogs += e.Message;
                return false;
            }

            Prog = new ComputeProgram(ComputeContext, sourceCode);

            try
            {
                Prog.Build(null, options, null, IntPtr.Zero);
            }
            catch (Exception e)
            {
                buildLogs = "Build failed.\r\n";
                buildLogs += e.Message;
                return false;
            }

            buildLogs = Prog.GetBuildLog(ComputeContext.Devices[deviceTypeIndex]);

            return true;
        }

        #region SetArguments

        /// <summary>
        /// Writes an array of type "Long" to the device.
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="values">Array of "Long".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        public bool SetMemoryArgument_Long(int argument_index, ref int[] values)
        {
            try
            {
                ComputeMemory varPointer = new ComputeBuffer<int>(ComputeContext, ComputeMemoryFlags.ReadWrite | ComputeMemoryFlags.CopyHostPointer, values);

                variablePointers[argument_index] = varPointer;
                kernel.SetMemoryArgument(argument_index, varPointer);

                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in SetMemoryArgument_Long: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// Writes an array of type "Single" to the device.
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="values">Array of "Single".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        public bool SetMemoryArgument_Single(int argument_index, ref float[] values)
        {
            try
            {
                ComputeMemory varPointer = new ComputeBuffer<float>(ComputeContext, ComputeMemoryFlags.ReadWrite | ComputeMemoryFlags.CopyHostPointer, values);

                variablePointers[argument_index] = varPointer;
                kernel.SetMemoryArgument(argument_index, varPointer);

                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in SetMemoryArgument_Single: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// Writes an array of type "Double" to the device.
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="values">Array of "Double".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        public bool SetMemoryArgument_Double(int argument_index, ref double[] values)
        {
            try
            {
                ComputeMemory varPointer = new ComputeBuffer<double>(ComputeContext, ComputeMemoryFlags.ReadWrite | ComputeMemoryFlags.CopyHostPointer, values);

                variablePointers[argument_index] = varPointer;
                kernel.SetMemoryArgument(argument_index, varPointer);

                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in SetMemoryArgument_Double: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// Sets "Long" argument to the kernel.
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="value_long">Argument value as "Long".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        public bool SetValueArgument_Long(int argument_index, int value_long)
        {
            try
            {
                ComputeMemory varPointer = new ComputeBuffer<int>(ComputeContext, ComputeMemoryFlags.ReadWrite, 1);

                variablePointers[argument_index] = varPointer;
                kernel.SetValueArgument(argument_index, value_long);

                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in SetValueArgument_Long: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// Sets "Single" argument to the kernel.
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="value_single">Argument value as "Single".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        public bool SetValueArgument_Single(int argument_index, float value_single)
        {
            try
            {
                ComputeMemory varPointer = new ComputeBuffer<float>(ComputeContext, ComputeMemoryFlags.ReadWrite, 1);

                variablePointers[argument_index] = varPointer;
                kernel.SetValueArgument(argument_index, value_single);

                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in SetValueArgument_Single: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// Sets "Double" argument to the kernel.
        /// </summary>
        /// <param name="argument_index">The argument index.</param>
        /// <param name="value_double">Argument value as "Double".</param>
        /// <returns>True, if the operation was successful, false otherwise.</returns>
        public bool SetValueArgument_Double(int argument_index, double value_double)
        {
            try
            {
                ComputeMemory varPointer = new ComputeBuffer<double>(ComputeContext, ComputeMemoryFlags.ReadWrite, 1);

                variablePointers[argument_index] = varPointer;
                kernel.SetValueArgument(argument_index, value_double);

                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in SetValueArgument_Double: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        #endregion SetArguments

        #region Execution

        /// <summary>
        /// Synchronous execution.
        /// </summary>
        /// <param name="globalWorkOffset">Array of global work offset, or "null".</param>
        /// <param name="globalWorkSize">Array of global work size, or "null".</param>
        /// <param name="localWorkSize">Array of local work size, or "null".</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        public bool ExecuteSync(ref int[] globalWorkOffset, ref int[] globalWorkSize, ref int[] localWorkSize)
        {
            try
            {
                ExecutionCompleted = false;
                InitGlobalArrays(ref globalWorkOffset, ref globalWorkSize, ref localWorkSize);

                ComputeCommandQueue.Execute(kernel, _globalWorkOffset, _globalWorkSize, _localWorkSize, null);
                ComputeCommandQueue.Finish();
                ExecutionCompleted = true;
                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in ExecuteSync: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// Execution in background.
        /// </summary>
        /// <param name="globalWorkOffset">Array of global work offset, or "null".</param>
        /// <param name="globalWorkSize">Array of global work size, or "null".</param>
        /// <param name="localWorkSize">Array of local work size, or "null".</param>
        /// <param name="threadPriority">Thread priority as integer (0 - "Lowest", 1 - "BelowNormal", 2 - "Normal", 3 - "AboveNormal", 4 - "Highest").</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        public bool ExecuteBackground(ref int[] globalWorkOffset, ref int[] globalWorkSize, ref int[] localWorkSize, int threadPriority)
        {
            try
            {
                ExecutionCompleted = false;
                InitGlobalArrays(ref globalWorkOffset, ref globalWorkSize, ref localWorkSize);

                if (threadPriority < (int)ThreadPriority.Lowest || threadPriority > (int)ThreadPriority.Highest)
                {
                    ErrorString += "\r\nError in ExecuteBackground: threadPriority = " + threadPriority + " is below 0 or above 4.";
                    return false;
                }

                Thread executionThread = new Thread(ExecutionThread)
                {
                    Name = "ExecutionThread",
                    Priority = (ThreadPriority)threadPriority
                };
                executionThread.Start();
                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in ExecuteBackground: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// For asynchronous call from VBA we need an address of callback function because VBA can use events only
        /// from form/classes.
        /// </summary>
        /// <param name="globalWorkOffset">Array of global work offset, or "null".</param>
        /// <param name="globalWorkSize">Array of global work size, or "null".</param>
        /// <param name="localWorkSize">Array of local work size, or "null".</param>
        /// <param name="threadPriority">Thread priority as integer (0 - "Lowest", 1 - "BelowNormal", 2 - "Normal", 3 - "AboveNormal", 4 - "Highest").</param>
        /// <param name="callback">Callback to the VBA function.</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        public bool ExecuteAsync(ref int[] globalWorkOffset, ref int[] globalWorkSize, ref int[] localWorkSize, int threadPriority,
            [MarshalAs(UnmanagedType.FunctionPtr)] ref Action callback)
        {
            try
            {
                this.callBack = callback;
                ExecuteBackground(ref globalWorkOffset, ref globalWorkSize, ref localWorkSize, threadPriority);
                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in ExecuteAsync: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// True, if execution is completed, false otherwise.
        /// </summary>
        public bool ExecutionCompleted { get; set; } = false;

        private void ExecutionThread()
        {
            ComputeCommandQueue.Execute(kernel, _globalWorkOffset, _globalWorkSize, _localWorkSize, null);
            ComputeCommandQueue.Finish();
            ExecutionCompleted = true;
            callBack?.Invoke();
        }

        private bool InitGlobalArrays(ref int[] globalWorkOffset, ref int[] globalWorkSize, ref int[] localWorkSize)
        {
            try
            {
                if (globalWorkOffset != null)
                {
                    _globalWorkOffset = new long[globalWorkOffset.Length];
                    globalWorkOffset.CopyTo(_globalWorkOffset, 0);
                }
                if (globalWorkSize != null)
                {
                    _globalWorkSize = new long[globalWorkSize.Length];
                    globalWorkSize.CopyTo(_globalWorkSize, 0);
                }
                if (localWorkSize != null)
                {
                    _localWorkSize = new long[localWorkSize.Length];
                    localWorkSize.CopyTo(_localWorkSize, 0);
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in InitGlobalArrays: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        #endregion Execution

        #region GetArguments

        /// <summary>
        /// Reads an array of type "Long" from the device.
        /// </summary>
        /// <param name="varIndex">0-based number of argument in argument list.</param>
        /// <param name="values">Array of "Long".</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        public bool GetMemoryArgument_Long(int varIndex, ref int[] values)
        {
            try
            {
                unsafe
                {
                    fixed (int* p = (int[])values)
                    {
                        IntPtr ptr = (IntPtr)p;
                        ComputeCommandQueue.Read((ComputeBuffer<int>)variablePointers[varIndex], true, 0L, values.Length, ptr, null);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in GetMemoryArgument_Long: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// Reads an array of type "Single" from the device.
        /// </summary>
        /// <param name="varIndex">0-based number of argument in argument list.</param>
        /// <param name="values">Array of "Single".</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        public bool GetMemoryArgument_Single(int varIndex, ref float[] values)
        {
            try
            {
                unsafe
                {
                    fixed (float* p = (float[])values)
                    {
                        IntPtr ptr = (IntPtr)p;
                        ComputeCommandQueue.Read((ComputeBuffer<float>)variablePointers[varIndex], true, 0L, values.Length, ptr, null);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in GetMemoryArgument_Single: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        /// <summary>
        /// Reads an array of type "Double" from the device.
        /// </summary>
        /// <param name="varIndex">0-based number of argument in argument list.</param>
        /// <param name="values">Array of "Double".</param>
        /// <returns>False in case of error/exception. Otherwise true.</returns>
        public bool GetMemoryArgument_Double(int varIndex, ref double[] values)
        {
            try
            {
                unsafe
                {
                    fixed (double* p = (double[])values)
                    {
                        IntPtr ptr = (IntPtr)p;
                        ComputeCommandQueue.Read((ComputeBuffer<double>)variablePointers[varIndex], true, 0L, values.Length, ptr, null);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ErrorString += "\r\nError in GetMemoryArgument_Double: " + ex.Message;
                ErrorString += "\r\n" + ex.StackTrace;
                return false;
            }
        }

        #endregion GetArguments

        /// <summary>
        /// Device type of initialized device ("GPU" / "CPU").
        /// </summary>
        public string DeviceType { get; set; } = "";

        /// <summary>
        /// Error string.
        /// </summary>
        public string ErrorString { get; set; } = "";
    }
}