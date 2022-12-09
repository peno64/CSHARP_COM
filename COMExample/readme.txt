This dll is a COM aware object. It can be used as a regular .NET assembly (without COM) 
and at the same time it can be used as a COM object for use by COM aware applications.

To create a .NET COM object, a regular Class Library project must be created just like a .NET assembly would be made.
Then in project properties, the following must be done:
- Application, Assembly Information...: "Make assembly COM-visible" must be checked
- Build, "Register for COM interop" must be checked
- Signing, "Sign the assembly" must be checked and below that "Choose a strong name key file" in the combo choose new to create a .snk file

Then add a class to define the COM interface.

In the namespace, define an interface:

    [Guid("7e309e4b-bf1c-409a-9ad1-56fbb3f445dc")] // put a unique GUID here! For example via https://www.guidgenerator.com/online-guid-generator.aspx
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IExample
    {
        [DispId(1)]
        void function1(arguments);

        [DispId(2)]
        void function2(arguments);

		.
		.
		.
    }

And then also define the actual class:

    [Guid("cb2bfb20-5e99-4329-a843-81feb7c49c7c")] // put a unique GUID here! For example via https://www.guidgenerator.com/online-guid-generator.aspx
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IExample))] // Use the interface defined above
    public class Example : IExample  // Use the interface defined above
    {
		public void function(arguments)
		{
			.
			.
			.
		}

		public void function2(arguments)
		{
			.
			.
			.
		}
	}

Note that the GUIDs must be unique over COM objects because COM uses this GUID to identify the interface.

Note that if arrays are used in arguments that they must always be passed by reference. 
This because COM requires that. Passing arrays by value is not supported by COM.
In .NET this means that the ref attribute must always be used for arrays.
null arrays are not allowed because COM also doesn't support this.

When used as .NET assembly, the dll must only be referenced in the using project (via references) and that's it, just like another regular assembly.

When used as COM object, the dll *must* be registered first. That is COM. No way to go beyound that.
Registration must be done on each computer where the dll will be used.
Note that when the dll is compiled in visual studio, registration is done automatically.
When this dll must be used on another computer the registration must be done manually.
This is done via the regasm.exe command line program:
RegAsm COMExample.dll /Codebase /tlb:COMExample.tlb
RegAsm is a command that is part of the .NET framework so it can also be found there on disk. 
The dll is compiled with the .NET 4.0 framework so the RegAsm command from this version must be used.
Note that there is a 32 bit version and a 64-bit version of this program. 
If the 64-bit version of the RegAsm program is used, 64-bit applications will be able to use the (64bit) COM object
If the 32-bit version of the RegAsm program is used, 32-bit applications will be able to use the (32bit) COM object
Note that most COM aware applications are 32-bit applications.
The dll can be registered both as 32-bit as 64-bit COM object.
For example under C:\Windows\Microsoft.NET\Framework\v4.0.30319 the 32-bit version of RegAsm can be found 
and under C:\Windows\Microsoft.NET\Framework64\v4.0.30319 the 64-bit version of RegAsm can be found.
Without registration, the first access to the COM dll will give an error, even if the dll is in the same folder as the application.

An example excel sheet is provided that uses the COM dll. Depending on having a 32-bit or a 64-bit office version you must at least register the corresponding version.
It is fine to register both as 32-bit and 64-bit.
To test the excel example, open the excel sheet, then press Alt F11 and position the cursor in the testcom sub in the module Module1 and then press F5. 
Messageboxes will be shown which show that values are changed.