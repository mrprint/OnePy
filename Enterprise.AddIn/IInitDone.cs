using System.Runtime.InteropServices;

namespace Enterprise.AddIn
{
    [Guid("AB634001-F13D-11d0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IInitDone
	{
		void Init([MarshalAs(UnmanagedType.IDispatch)] object pConnection);
		void Done();
        void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info);
	}
}