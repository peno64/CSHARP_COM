using System;
using System.Runtime.InteropServices;

namespace COMExample
{
    [Guid("7e309e4b-bf1c-409a-9ad1-56fbb3f445dc")] // put a unique GUID here! For example via https://www.guidgenerator.com/online-guid-generator.aspx
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IExample
    {
        [DispId(1)]
        void function1(int a, string b, double c, bool d, ref int e);

        [DispId(2)]
        int function2(int a);
    }

    [Guid("cb2bfb20-5e99-4329-a843-81feb7c49c7c")] // put a unique GUID here! For example via https://www.guidgenerator.com/online-guid-generator.aspx
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IExample))] // Use the interface defined above
    public class Example : IExample  // Use the interface defined above
    {
        public void function1(int a, string b, double c, bool d, ref int e)
        {
            e = a * 2;
		}

        public int function2(int a)
        {
            return a * 2;
		}
    }
}