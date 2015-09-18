using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Collections;
using Microsoft.Scripting.Hosting;
using Microsoft.Scripting.Runtime;
using Microsoft.Scripting;
using IronPython.Compiler;
using IronPython.Runtime;
using System.Runtime.InteropServices.ComTypes;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using Enterprise.AddIn;

namespace OnePy
{
    
    [ComVisible(false)]
    public class Interactor
    {
        public object like_dispatch;
        public IAsyncEvent like_asyncevent;
        public IErrorLog like_errorlog;
        public IStatusLine like_statusline;
        public IExtWndsSupport like_extwndssupport;
        public IPropertyProfile like_propertyprofile;

        public Interactor()
        {
            like_dispatch = V7Data.V7Object;
            like_asyncevent = V7Data.AsyncEvent;
            like_errorlog = V7Data.ErrorLog;
            like_statusline = V7Data.StatusLine;
            like_extwndssupport = V7Data.ExtWndsSupport;
            like_propertyprofile = V7Data.PropertyProfile;
        }


        //~Interactor()
        //{
        //}

        public void Message(EXCEPINFOCover eic)
        {
            V7Data.ErrorLog.AddError("", ref eic.data);
            //throw new System.Exception("message");
        }

        #region EXCEPINFOCover
        [ComVisible(false)]
        public class EXCEPINFOCover // Наследовать EXCEPINFO
        {
            public EXCEPINFOCover()
            {
                data = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
            }

            public EXCEPINFOCover(System.Runtime.InteropServices.ComTypes.EXCEPINFO ei)
            {
                data = ei;
            }

            public static implicit operator System.Runtime.InteropServices.ComTypes.EXCEPINFO(EXCEPINFOCover eic)
            {
                return eic.data;
            }

            public short wCode
            {
                get
                {
                    return data.wCode;
                }
                set
                {
                    data.wCode = value;
                }
            }

            public short wReserved
            {
                get
                {
                    return data.wReserved;
                }
                set
                {
                    data.wReserved = value;
                }
            }

            public string bstrSource
            {
                get
                {
                    return data.bstrSource;
                }
                set
                {
                    data.bstrSource = value;
                }
            }

            public string bstrDescription
            {
                get
                {
                    return data.bstrDescription;
                }
                set
                {
                    data.bstrDescription = value;
                }
            }

            public string bstrHelpFile
            {
                get
                {
                    return data.bstrHelpFile;
                }
                set
                {
                    data.bstrHelpFile = value;
                }
            }

            public int dwHelpContext
            {
                get
                {
                    return data.dwHelpContext;
                }
                set
                {
                    data.dwHelpContext = value;
                }
            }

            public System.IntPtr pvReserved
            {
                get
                {
                    return data.pvReserved;
                }
                set
                {
                    data.pvReserved = value;
                }
            }

            public int scode
            {
                get
                {
                    return data.scode;
                }
                set
                {
                    data.scode = value;
                }
            }

            public System.Runtime.InteropServices.ComTypes.EXCEPINFO data;
        }
        #endregion
    }

#if (ONEPY45)
    public partial class OnePy45
#else
#if (ONEPY35)
    public partial class OnePy35
#endif
#endif
    {
        Interactor interactor;
        ScriptRuntime sruntime;
        ScriptScope conn_scope;
        ScriptScope interf_scope;
        ObjectOperations ops;

        #region Ссылки на методы питон-интерфейса
        Func<Object, Object> prGetNProps;
        Func<Object, Object, Object> prFindProp;
        Func<Object, Object, Object, Object> prGetPropName;
        Func<Object, Object, Object> prGetPropVal;
        Func<Object, Object, Object> prSetPropVal;
        Func<Object, Object, Object> prIsPropReadable;
        Func<Object, Object, Object> prIsPropWritable;
        Func<Object, Object> prGetNMethods;
        Func<Object, Object, Object> prFindMethod;
        Func<Object, Object, Object, Object> prGetMethodName;
        Func<Object, Object, Object> prGetNParams;
        Func<Object, Object, Object, Object> prGetParamDefValue;
        Func<Object, Object, Object> prHasRetVal;
        Func<Object, Object, Object> prCallAsProc;
        Func<Object, Object, Object, Object> prCallAsFunc;
        Func<Object> prCleanAll;
        Func<Object, Object> prAddReferences;
        Func<Object> prClearExcInfo;
        #endregion

        void InitDLR()
        {
            try
            {
                // Возможно app.config придется упразднять
                // ?собрать IronPython1C
                
#if (ONEPY45)
                ScriptRuntimeSetup setup = ScriptRuntimeSetup.ReadConfiguration(Path.Combine(AssemblyDirectory, "OnePy45.dll.config"));
#else
#if (ONEPY35)
                ScriptRuntimeSetup setup = ScriptRuntimeSetup.ReadConfiguration(Path.Combine(AssemblyDirectory, "OnePy35.dll.config"));
#endif
#endif
                //setup.Options.Add("PreferComDispatch", ScriptingRuntimeHelpers.True);
                sruntime = new ScriptRuntime(setup);
                ScriptEngine eng = sruntime.GetEngine("Python");
                
                #region Установка путей поиска модулей
                var sp = eng.GetSearchPaths();
                sp.Add(Environment.CurrentDirectory);
                sp.Add(Path.Combine(Environment.CurrentDirectory, @"Lib"));
                sp.Add(Path.Combine(Environment.CurrentDirectory, @"Lib\site-packages"));
                sp.Add(Path.Combine(Environment.CurrentDirectory, @"IronPython.lib"));
                sp.Add(Path.Combine(Environment.CurrentDirectory, @"IronPython"));
                sp.Add(Path.Combine(Environment.CurrentDirectory, @"IronPython\DLLs"));
                sp.Add(Path.Combine(Environment.CurrentDirectory, @"IronPython\Lib"));
                sp.Add(Path.Combine(Environment.CurrentDirectory, @"IronPython\Lib\site-packages"));
                sp.Add(Path.Combine(AssemblyDirectory, @"IronPython.lib"));
                sp.Add(Path.Combine(AssemblyDirectory, @"IronPython"));
                sp.Add(Path.Combine(AssemblyDirectory, @"IronPython\DLLs"));
                sp.Add(Path.Combine(AssemblyDirectory, @"IronPython\Lib"));
                sp.Add(Path.Combine(AssemblyDirectory, @"IronPython\Lib\site-packages"));
                
                foreach (string ap in RuntimeConfig.additional_paths)
                {
                    sp.Add(ap);
                }
                
                sp.Add(AssemblyDirectory);
                eng.SetSearchPaths(sp);
                #endregion
                
                ScriptSource conn_src = eng.CreateScriptSource(new AssemblyStreamContentProvider("OnePy.#1"), "cm5ACF5D43F2DA488BB5414714845ACBDE.py");
                ScriptSource interf_src = eng.CreateScriptSource(new AssemblyStreamContentProvider("OnePy.#2"), "interfacing.py");
                
                var comp_options = (PythonCompilerOptions)eng.GetCompilerOptions();
                comp_options.Optimized = false;
                comp_options.Module &= ~ModuleOptions.Optimized; 
                interf_scope = eng.CreateScope();
                interf_src.Compile(comp_options).Execute(interf_scope);
                conn_scope = eng.CreateScope();
                conn_src.Compile(comp_options).Execute(conn_scope);
                ops = eng.CreateOperations();
                Object calcClass = conn_scope.GetVariable("OnePyConnector");
                interactor = new Interactor();
                Object calcObj = ops.Invoke(calcClass, interactor);
                
                #region Получение ссылок на методы
                prGetNProps = ops.GetMember<Func<Object, Object>>(calcObj, "GetNProps");
                prFindProp = ops.GetMember<Func<Object, Object, Object>>(calcObj, "FindProp");
                prGetPropName = ops.GetMember<Func<Object, Object, Object, Object>>(calcObj, "GetPropName");
                prGetPropVal = ops.GetMember<Func<Object, Object, Object>>(calcObj, "GetPropVal");
                prSetPropVal = ops.GetMember<Func<Object, Object, Object>>(calcObj, "SetPropVal");
                prIsPropReadable = ops.GetMember<Func<Object, Object, Object>>(calcObj, "IsPropReadable");
                prIsPropWritable = ops.GetMember<Func<Object, Object, Object>>(calcObj, "IsPropWritable");
                prGetNMethods = ops.GetMember<Func<Object, Object>>(calcObj, "GetNMethods");
                prFindMethod = ops.GetMember<Func<Object, Object, Object>>(calcObj, "FindMethod");
                prGetMethodName = ops.GetMember<Func<Object, Object, Object, Object>>(calcObj, "GetMethodName");
                prGetNParams = ops.GetMember<Func<Object, Object, Object>>(calcObj, "GetNParams");
                prGetParamDefValue = ops.GetMember<Func<Object, Object, Object, Object>>(calcObj, "GetParamDefValue");
                prHasRetVal = ops.GetMember<Func<Object, Object, Object>>(calcObj, "HasRetVal");
                prCallAsProc = ops.GetMember<Func<Object, Object, Object>>(calcObj, "CallAsProc");
                prCallAsFunc = ops.GetMember<Func<Object, Object, Object, Object>>(calcObj, "CallAsFunc");
                prCleanAll = ops.GetMember<Func<Object>>(calcObj, "CleanAll");
                prAddReferences = ops.GetMember<Func<Object, Object>>(calcObj, "AddReferences");
                prClearExcInfo = ops.GetMember<Func<Object>>(calcObj, "ClearExcInfo");
                #endregion

            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }

        }

        #region "Свойства"

        private void PyGetNProps(ref int plProps)
		{
            try
            {
                plProps = (int)prGetNProps(plProps);
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyFindProp(string bstrPropName, ref int plPropNum)
		{
            try
            {
                plPropNum = (int)prFindProp(bstrPropName, plPropNum);
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyGetPropName(int lPropNum, int lPropAlias, ref string pbstrPropName)
		{
            try
            {
                pbstrPropName = (string)prGetPropName(lPropNum, lPropAlias, pbstrPropName);
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyGetPropVal(int lPropNum, ref object pvarPropVal)
		{
            try
            {
                pvarPropVal = prGetPropVal(lPropNum, pvarPropVal);
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PySetPropVal(int lPropNum, ref object varPropVal)
		{
            try
            {
                varPropVal = prSetPropVal(lPropNum, varPropVal);
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyIsPropReadable(int lPropNum, ref bool pboolPropRead)
		{
            try
            {
			    pboolPropRead = (bool)prIsPropReadable(lPropNum, pboolPropRead); // Все свойства доступны для чтения
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyIsPropWritable(int lPropNum, ref bool pboolPropWrite)
		{
            try
            {
                pboolPropWrite = (bool)prIsPropWritable(lPropNum, pboolPropWrite); // Все свойства доступны для записи
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}
		#endregion
	
		#region "Методы"

        private void PyGetNMethods(ref int plMethods)
		{
            try
            {
                plMethods = (int)prGetNMethods(plMethods);
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyFindMethod(string bstrMethodName, ref int plMethodNum)
		{
            try
            {
			    plMethodNum = (int)prFindMethod(bstrMethodName, plMethodNum);
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyGetMethodName(int lMethodNum, int lMethodAlias, ref string pbstrMethodName)
		{
            try
            {
			    pbstrMethodName = (string)prGetMethodName(lMethodNum, lMethodAlias, pbstrMethodName);
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyGetNParams(int lMethodNum, ref int plParams)
		{
            try
            {
                plParams = (int)prGetNParams(lMethodNum, plParams);
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyGetParamDefValue(int lMethodNum, int lParamNum, ref object pvarParamDefValue)
		{
            try
            {
			    pvarParamDefValue = prGetParamDefValue(lMethodNum, lParamNum, pvarParamDefValue); //Нет значений по умолчанию
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyHasRetVal(int lMethodNum, ref bool pboolRetValue)
		{
            try
            {
			    pboolRetValue = (bool)prHasRetVal(lMethodNum, pboolRetValue);  //Все методы у нас будут функциями (т.е. будут возвращать значение). 
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyCallAsProc(int lMethodNum, ref object[] pParams)
		{
            try
            {
                //pParams = (object[])
                prCallAsProc(lMethodNum, pParams);
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

        private void PyCallAsFunc(int lMethodNum, ref object pvarRetValue, ref object[] pParams)
		{
            try
            {
			    pvarRetValue = prCallAsFunc(lMethodNum, pvarRetValue, pParams); //Возвращаемое значение метода для 1С			
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
		}

		#endregion

        void CleanAll()
        {
            try
            {
                prCleanAll();
            }
            catch (Exception e)
            {
                ProcessError(e);
                throw;
            }
        }

        void FreePython()
        {
            prGetNProps = null;
            prFindProp = null;
            prGetPropName = null;
            prGetPropVal = null;
            prSetPropVal = null;
            prIsPropReadable = null;
            prIsPropWritable = null;
            prGetNMethods = null;
            prFindMethod = null;
            prGetMethodName = null;
            prGetNParams = null;
            prGetParamDefValue = null;
            prHasRetVal = null;
            prCallAsProc = null;
            prCallAsFunc = null;
            prCleanAll = null;
            prAddReferences = null;
            prClearExcInfo = null;
            
            ops = null;
            conn_scope = null;
            interf_scope = null;
            interactor = null;
            sruntime.Shutdown();
        }
    }
}