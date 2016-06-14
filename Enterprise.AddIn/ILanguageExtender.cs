using System.Runtime.InteropServices;

namespace Enterprise.AddIn
{
    [Guid("AB634003-F13D-11d0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ILanguageExtender
	{
        void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref string bstrExtensionName);
		void GetNProps(ref int plProps);
        void FindProp([MarshalAs(UnmanagedType.BStr)] string bstrPropName, ref int plPropNum);
        void GetPropName(int lPropNum, int lPropAlias, [MarshalAs(UnmanagedType.BStr)] ref string pbstrPropName);
        void GetPropVal(int lPropNum, [MarshalAs(UnmanagedType.Struct)] ref object pvarPropVal);
        void SetPropVal(int lPropNum, [MarshalAs(UnmanagedType.Struct)] ref object varPropVal);
		void IsPropReadable(int lPropNum, ref bool pboolPropRead);
		void IsPropWritable(int lPropNum, ref bool pboolPropWrite);
		void GetNMethods(ref int plMethods);
        void FindMethod([MarshalAs(UnmanagedType.BStr)] string bstrMethodName, ref int plMethodNum);
        void GetMethodName(int lMethodNum, int lMethodAlias, [MarshalAs(UnmanagedType.BStr)] ref string pbstrMethodName);
		void GetNParams(int lMethodNum, ref int plParams);
        void GetParamDefValue(int lMethodNum, int lParamNum, [MarshalAs(UnmanagedType.Struct)] ref object pvarParamDefValue);
		void HasRetVal(int lMethodNum, ref bool pboolRetValue);
        void CallAsProc(int lMethodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams);
        void CallAsFunc(int lMethodNum, [MarshalAs(UnmanagedType.Struct)] ref object pvarRetValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParam);
	}
}