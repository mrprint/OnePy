using System;
using System.Runtime.InteropServices;

namespace Enterprise.AddIn
{
    [ComImport, Guid("55272A00-42CB-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyBag
    {
        void Write([InAttribute] string propName, [InAttribute] ref Object ptrVar);
        void Read([InAttribute] string propName, out Object ptrVar, int errorLog);
    }

    //[ComImport, Guid("37D84F60-42CB-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    //interface IPersistPropertyBag
    //{

    //    [PreserveSig]
    //    void InitNew();

    //    [PreserveSig]
    //    void Load(IPropertyBag propertyBag, int errorLog);

    //    [PreserveSig]
    //    void Save(IPropertyBag propertyBag, [InAttribute] bool clearDirty, [InAttribute] bool saveAllProperties);

    //    [PreserveSig]
    //    void GetClassID(out Guid classID);
    //}

    [Guid("AB634002-F13D-11d0-A459-004095E1DAEA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyProfile : IPropertyBag
    {
        [PreserveSig]
        uint RegisterProfileAs([MarshalAs(UnmanagedType.BStr)] string bstrProfileName);
    }
}