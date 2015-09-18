using System;
using System.Runtime.InteropServices;

namespace Enterprise.AddIn
{
    [Guid("00020400-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IDispatch
    {
        [PreserveSig]
        uint GetTypeInfoCount(out int Count);
        [PreserveSig]
        uint GetTypeInfo([MarshalAs(UnmanagedType.U4)] int iTInfo,
                        [MarshalAs(UnmanagedType.U4)] int lcid, out IntPtr ppTinfo); //? ITypeInfo
        [PreserveSig]
        uint GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] 
                          string[] rgsNames, int cNames, int lcid, [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig]
        uint Invoke(int dispIdMember, ref Guid riid, uint lcid, ushort wFlags, 
                   ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, out object pVarResult,
                   ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, IntPtr[] pArgErr);
    }
}