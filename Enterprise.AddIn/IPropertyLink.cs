using System;
using System.Runtime.InteropServices;

namespace Enterprise.AddIn
{
    [Guid("52512A61-2A9D-11d1-A4D6-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface IPropertyLink
	{
        void get_Enabled(ref char[] pData, int Id);
        void put_Enabled(ref char[] pData, int Id);
	}
}