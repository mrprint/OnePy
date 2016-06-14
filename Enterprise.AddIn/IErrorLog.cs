using System.Runtime.InteropServices;

namespace Enterprise.AddIn
{
    [Guid("3127CA40-446E-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IErrorLog
    {
        [PreserveSig]
        uint AddError([MarshalAs(UnmanagedType.BStr)] string pszPropName, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExepInfo);
    }
}