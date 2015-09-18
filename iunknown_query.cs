public ITypeInfo GetCOMObjectsTypeInfo(object rcw/*IntPtr com*/)
        {
            //var rcw = Marshal.GetObjectForIUnknown(com);
            //Marshal.Release(com);
            var idisp = (Enterprise.AddIn.IDispatch)rcw;
            int count = 0;
            idisp.GetTypeInfoCount(out count);
            if (count < 1)
                throw new System.Exception("No type info");
            IntPtr _typeinfo;
            idisp.GetTypeInfo(0, 0, out _typeinfo);
            if (_typeinfo == null)
                throw new System.Exception("No type info");
            var typeInfo = (ITypeInfo)Marshal.GetTypedObjectForIUnknown(_typeinfo, typeof(ITypeInfo));
            Marshal.Release(_typeinfo);
            return typeInfo;
        }

        public void DumpTypeInfo(ITypeInfo typeInfo)
        {
            string typeName = "";
            string DocString;
            int HelpContext;
            string HelpFile;
            typeInfo.GetDocumentation(-1, out typeName, out DocString, out HelpContext, out HelpFile);
            //Console.WriteLine("TYPE {0}", typeName);
            Log.DWrite(String.Format("TYPE {0}", typeName));
            IntPtr pTypeAttr;
            typeInfo.GetTypeAttr(out pTypeAttr);
            ComTypes.TYPEATTR typeAttr = (ComTypes.TYPEATTR)Marshal.PtrToStructure(pTypeAttr, typeof(ComTypes.TYPEATTR));
            typeInfo.ReleaseTypeAttr(pTypeAttr);
            for (int iImplType = 0; iImplType < typeAttr.cImplTypes; ++iImplType)
            {
                int href;
                typeInfo.GetRefTypeOfImplType(iImplType, out href);
                ITypeInfo implTypeInfo;
                typeInfo.GetRefTypeInfo(href, out implTypeInfo);
                string implTypeName = "";
                implTypeInfo.GetDocumentation(-1, out implTypeName, out DocString, out HelpContext, out HelpFile);
                //Console.WriteLine("  Implements {0}", implTypeName);
                Log.DWrite(String.Format("  Implements {0}", implTypeName));
            }
            for (int iVar = 0; iVar < typeAttr.cVars; ++iVar)
            {
                IntPtr pVarDesc;
                typeInfo.GetVarDesc(iVar, out pVarDesc);
                ComTypes.VARDESC varDesc = (ComTypes.VARDESC)Marshal.PtrToStructure(pVarDesc, typeof(ComTypes.VARDESC));
                typeInfo.ReleaseVarDesc(pVarDesc);
                string[] names = { "" };
                int pcNames;
                typeInfo.GetNames(varDesc.memid, names, 1, out pcNames);
                var varName = names[0];
                //Console.WriteLine("  Dim {0} As {1}", varName, DumpTypeDesc(varDesc.elemdescVar.tdesc, typeInfo));
                Log.DWrite(string.Format("  Dim {0} As {1}", varName, DumpTypeDesc(varDesc.elemdescVar.tdesc, typeInfo)));
            }
        }

        string DumpTypeDesc(ComTypes.TYPEDESC tdesc, ComTypes.ITypeInfo context)
        {
            VarEnum vt = (VarEnum)tdesc.vt;
            ComTypes.TYPEDESC tdesc2;
            switch (vt)
            {
                case VarEnum.VT_PTR:
                    tdesc2 = (ComTypes.TYPEDESC)Marshal.PtrToStructure(tdesc.lpValue, typeof(ComTypes.TYPEDESC));
                    return "Ref " + DumpTypeDesc(tdesc2, context);
                case VarEnum.VT_USERDEFINED:
                    int href = (int)(tdesc.lpValue.ToInt64() & int.MaxValue);
                    ComTypes.ITypeInfo refTypeInfo = null;
                    context.GetRefTypeInfo(href, out refTypeInfo);
                    string typeName = "";
                    string DocString;
                    int HelpContext;
                    string HelpFile;
                    refTypeInfo.GetDocumentation(-1, out typeName, out DocString, out HelpContext, out HelpFile);
                    return typeName;
                case VarEnum.VT_CARRAY:
                    tdesc2 = (ComTypes.TYPEDESC)Marshal.PtrToStructure(tdesc.lpValue, typeof(ComTypes.TYPEDESC));
                    return "Array of " + DumpTypeDesc(tdesc2, context);
                    // lpValue is actually an ARRAYDESC structure, which also has
                    // information on the array dimensions, but alas .NET doesn't 
                    // predefine ARRAYDESC.
                default:
                    // There are many other VT_s that I haven't special-cased, 
                    // e.g. VT_INTEGER.
                    return vt.ToString();
            }
        }