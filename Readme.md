# COM-wrapper of [Cloo](https://github.com/clSharp/Cloo) to execute OpenCL code from Excel.
The wrapper allows to execute OpenCL code on CPU and GPU devices from VBA.
More detailed description with examples can be found in [my CodeProject article](https://www.codeproject.com/Articles/5332060/How-to-Use-GPU-in-VBA-Excel).

The wrapper has simple implementation and divided in two independent parts:
- <p style='text-align:justify'>ClooWrapperVBA.Configuration, to obtain configuration of available platforms and associated CPUs and GPUs.</p>
- <p style='text-align:justify'>ClooWrapperVBA.ProgramDevice, to compile and start OpenCL programs on CPUs and GPUs and obtain the results. It is also possible to start programs on CPUs and GPUs simultaneously (asynchronously). In asynchronous mode it is also possible to set the priority of execution.</p>
<br>

## Downloads
The current version can be downloaded as:
* [Installer for Windows](https://sourceforge.net/projects/cloowrappervba/files/ClooWrapperVBA%20setup.exe/download). Installation path is "C:\Program Files (x86)\ClooWrapperVBA\".
* [Zip-file which contains the same content as the installer](https://sourceforge.net/projects/cloowrappervba/files/ClooWrapperVBA.zip/download). The components must be registered using "register.bat". Please note, that the "register.bat" must be started with admin rights.

<p style='text-align:justify'>Directory "demo" contains the Excel table with all demos and VBscript file to check available platforms and devices even without Excel.</p>
<br>

## Dependencies
* .Net framework version 4.0.
<br>

## Available functions.
* ClooWrapperVBA.Configuration:
    * <p style='text-align:justify'>ClooWrapperVBA.Configuration.Platform - contains information on the available platform.</p>
    * <p style='text-align:justify'>ClooWrapperVBA.Configuration.Platform.Device - contains information on each device in available platforms. In total you can obtain 59 device-specific properties.</p>


* ClooWrapperVBA.ProgramDevice:
    * <p style='text-align:justify'>ClooWrapperVBA.ProgramDevice.Build - compiles sources for selected device.</p>
    * <p style='text-align:justify'>ClooWrapperVBA.ProgramDevice.CreateKernel - Loads the function to execute.</p>
    * <p style='text-align:justify'>ClooWrapperVBA.ProgramDevice.SetValueArgument_..., ClooWrapperVBA.ProgramDevice.SetMemoryArgument_... - Sets argument values and arrays of integers, floats and doubles of the function to execute. The parameter "argument_index" starts with 0 for first argument and must be manually incrased for the next arguments. It is also very important to set variables in a right sequence. First, the variable with argument index 0, then with argument index 1 and so on.</p>
    * <p style='text-align:justify'>ClooWrapperVBA.ProgramDevice.ExecuteSync - Execute function synchronously. Excel will move further only after execution was completed.</p>
    * <p style='text-align:justify'>ClooWrapperVBA.ProgramDevice.ExecuteAsync - Start execution of the function asynchronously. The callback function will be called at the end of execution.</p>
    * <p style='text-align:justify'>ClooWrapperVBA.ProgramDevice.ExecuteBackground - Start execution of the function asynchronously. After execution the flag "ClooWrapperVBA.ProgramDevice.ExecutionCompleted" is set to true.</p>
    * <p style='text-align:justify'>ClooWrapperVBA.ProgramDevice.GetMemoryArgument_... - Read arguments (results) from the function.</p>
    * <p style='text-align:justify'>ClooWrapperVBA.ProgramDevice.ReleaseMemObject - Releases instantiated memory objects. The single parameter has the same meaning as "argument_index" from SetValue/MemoryArguments. It should start with highest used "argument_index".</p>
    * <p style='text-align:justify'>ClooWrapperVBA.ProgramDevice.ReleaseKernel and ClooWrapperVBA.ProgramDevice.ReleaseProgram do the rest of disposing of instantiated OpenCL parts.</p>
<br>

## VBA samples.
1. "Configuration": Prints configuration of all platforms and available devices.
2. "Performance":
    - "VBA performance test" uses multiplication of two 1200x1200 matrices and measures execution times. Also, the correctness of results of matrix multiplications on CPU/GPU is compared to VBA results (column "C").
    - The CL code of performance measurements is taken from a CodeProject article "[How to Use Your GPU in .NET](https://www.codeproject.com/Articles/1116907/How-to-Use-Your-GPU-in-NET)".
3. <p style='text-align:justify'>"Asynchronous": For asynchronous execution example a matrix multiplication of two 2000x2000 matrices is used.</p>
<br>

## VBscript sample.
**Good news**: Yes! It is also possible to use the wrapper also from VBscript!
<br>
**Bad news**: You can obtain only configuration of the platforms and devices. The reason: VBscript uses variants for arrays. Of course it is possible to use object or ArrayList to set the arrays, but this will make the wrapper much complicated and is out of the scope of the wrapper.
<br>

## Helpful VBA functions.
- CTimer class is taken from the article "[How do you test running time of VBA code?](https://stackoverflow.com/questions/198409/how-do-you-test-running-time-of-vba-code)". It implements a very precise timer to measure performance.
- <p style='text-align:justify'>"MatrixToVector" and "VectorToMatrix" are used as expected from their names to load arrays to CPU/GPU and get the results of execution back to VBA in matrix form.</p>
<br>

## Implementation notes:
1. Cloo version 0.9.1.0 is used. The reason: the wrapper was intended to work with version 4.0 of .Net framework.
2. **Build**: Parameters "deviceIndex", "deviceTypeIndex" have different meaning. 
    - <p style='text-align:justify'>"deviceIndex" is a device index of devices at platform defined in "platformIndex". The devices corresponding to each "deviceIndex" can be of different type ("CPU"/"GPU").</p> 
    - "deviceTypeIndex" is an index inside of same device type. 
    - <p style='text-align:justify'>Example: If your platform have 3 devices, one CPU and two GPUs, then the possible "deviceIndex" values are 0 (GPU), 1 (CPU) and 2 (GPU). The "deviceTypeIndex" in this configuration will be 0 and 1 for GPUs and 0 for CPU. You can obtain the sequence of devices using the Configuration.</p>
    - <p style='text-align:justify'>To simplify usage from VBA, all devices can be added to the collection using function "CreateDeviceCollection". You can obtain the first CPU and first GPU using a function "GetFirstDeviceOfType" where the first argument is a collection of devices and the second argument is a device type, "CPU" or "GPU". The collection of all available devices is also very useful to run your code in asynchronous mode at all available devices.</p>
3. **Build**: Parameter "options" contain compiler options. In simplest case it can be empty ("", not "null" or "Nothing"). Among the common compiler oprions, like "-w" (inhibit all warning messages), you can also define here commonly used constants ("-D name=definition") and use them in the OpenCL code. The complete list of compiler options can be found at [official Khronos home page](https://www.khronos.org/registry/OpenCL/sdk/1.0/docs/man/xhtml/clBuildProgram.html).
4. **ExecuteAsync** function must be used with care:
    - <p style='text-align:justify'>During debugging Excel can crash because of simultaneous execution of the code in callback and "MainLoop" functions.</p>
    - <p style='text-align:justify'>Writing out of the results to the cells in callback function can also cause an Excel crash.</p>
    - A good solution is to use instead **ExecuteBackground** function.
5. ReleaseMemObject, ReleaseKernel and ReleaseProgram are added to accurately dispose instantiated OpenCL objects and to avoid side effects from not disposed objects. Nevertheless, the current code in Excel example works correctly also without them.
<br>

## Not tested parts:
- <p style='text-align:justify'>globalWorkOffset, localWorkSize were not tested and were added analogously to globalWorkSize.</p>

## FAQ:
1. Configuration: No platforms/devices were found.
    - Reason: OpenCL.dll is not found in "Windows" folder.
        - Solution: Get OpenCL.dll from other computer.
    - Reason: The GPGPU / CPU drivers are too old and not supported by OpenCL.
        - Update the drivers.
<br>