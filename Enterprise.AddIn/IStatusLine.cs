using System.Runtime.InteropServices;

namespace Enterprise.AddIn
{
	[Guid("ab634005-f13d-11d0-a459-004095e1daea"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IStatusLine
	{
        [PreserveSig]
		uint SetStatusLine([MarshalAs(UnmanagedType.BStr)] string bstrStatusLine);
        [PreserveSig]
		uint ResetStatusLine();
	}
}