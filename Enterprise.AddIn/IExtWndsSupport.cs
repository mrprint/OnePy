using System;
using System.Runtime.InteropServices;

namespace Enterprise.AddIn
{
    [Guid("efe19ea0-09e4-11d2-a601-008048da00de"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IExtWndsSupport
    {
        [PreserveSig]
        uint GetAppMainFrame(ref IntPtr hwnd);
        [PreserveSig]
        uint GetAppMDIFrame(ref IntPtr hwnd);
        [PreserveSig]
        uint CreateAddInWindow(
            [MarshalAs(UnmanagedType.BStr)] string bstrProgID,
            [MarshalAs(UnmanagedType.BStr)] string bstrWindowName,
            uint dwStyles,
            uint dwExStyles,
            ref RECT rctl,
            uint Flags,
            ref IntPtr pHwnd,
            [MarshalAs(UnmanagedType.IDispatch)] ref object[] pDisp);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }
}