using System;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace Pptcom.ComProxy
{
    // .NET Framework does not publicly expose IDispatch - declare just enough of it to call
    // GetTypeInfo, matching the well-known IID for IDispatch.
    [ComImport]
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDispatchMinimal
    {
        void GetTypeInfoCount(out int count);

        void GetTypeInfo(int index, int lcid, out ComTypes.ITypeInfo typeInfo);

        // GetIDsOfNames / Invoke omitted - not needed here, and leaving them out keeps the
        // vtable-order-sensitive parts of this interface as small as possible to get right.
    }
}
