using System.Runtime.InteropServices;

namespace Enterprise.AddIn
{
    [Guid("ab634004-f13d-11d0-a459-004095e1daea"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IAsyncEvent
	{
        [PreserveSig]
        uint SetEventBufferDepth(int lDepth);
        [PreserveSig]
        uint GetEventBufferDepth(ref int plDepth);
        [PreserveSig]
        uint ExternalEvent(
            [MarshalAs(UnmanagedType.BStr)] string bstrSource, 
            [MarshalAs(UnmanagedType.BStr)] string bstrMessage,
            [MarshalAs(UnmanagedType.BStr)] string bstrData);
        [PreserveSig]
        uint CleanBuffer();
	}
}
