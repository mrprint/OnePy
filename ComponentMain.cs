using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Enterprise.AddIn;


namespace OnePy
{
#if (ONEPY45)
    [ComVisible(true), Guid("7653664F-920A-48AF-AE16-7662BF963886"), ProgId("AddIn.OnePy45")]
    public partial class OnePy45 : IInitDone, ILanguageExtender
#else
#if (ONEPY35)
    [ComVisible(true), Guid("F0416ABF-AED5-400A-AC7B-13022717D466"), ProgId("AddIn.OnePy35")]
    public partial class OnePy35 : IInitDone, ILanguageExtender
#endif
#endif
    {
#if (ONEPY45)
        const string c_AddinName = "OnePy45";
#else
#if (ONEPY35)
        const string c_AddinName = "OnePy35";
#endif
#endif
        private static int InstancesCount = 0;
        private static AppDomain domain;
        private bool Initialized;

#if (ONEPY45)
        public OnePy45()
#else
#if (ONEPY35)
        public OnePy35()
#endif
#endif
        {
            //domain = AppDomain.CreateDomain("OnePy", AppDomain.CurrentDomain.Evidence);
            Initialized = false;
            Log.PrepareDir();
        }

#if (ONEPY45)
        ~OnePy45()
#else
#if (ONEPY35)
        ~OnePy35()
#endif
#endif
        {
            DoneAll();
            //AppDomain.Unload(domain);
        }

        void DoneAll()
        {
            if (Initialized)
            {
                CleanAll();
                FreePython();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            Initialized = false;
        }

        void InitByNeed()
        {
            if (!Initialized)
            {
                try
                {
                    LoadConfiguration();
                    InitDLR();
                    prAddReferences(RuntimeConfig.additional_asms);
                    Initialized = true;
                }
                catch (Exception e)
                {
                    ProcessError(e);
                    throw;
                }
            }
        }

        #region "IInitDone implementation"

        public void Init([MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            if (InstancesCount < 1)
            {
                V7Data.V7Object = pConnection;
            }
            InstancesCount += 1;
        }

        public void Done()
        {
            InstancesCount -= 1;
            DoneAll();
            if (InstancesCount < 1)
            {
                V7Data.Clean();
            }
        }

        public void GetInfo([MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] info)
        {
            info[0] = 2000;
        }

        #endregion

        public void RegisterExtensionAs([MarshalAs(UnmanagedType.BStr)] ref string bstrExtensionName)
        {
            bstrExtensionName = c_AddinName;
        }

        #region "Свойства"

        enum Props
        {   //Числовые идентификаторы свойств нашей внешней компоненты
            //propMessageBoxIcon = 0,  //Пиктограмма для MessageBox'а
            //propMessageBoxButtons = 1, //Кнопки для MessageBox'a
            LastProp = 0
        }

        public void GetNProps(ref int plProps)
        {	//Здесь 1С получает количество доступных из ВК свойств
            InitByNeed();
            PyGetNProps(ref plProps);
        }

        public void FindProp([MarshalAs(UnmanagedType.BStr)] string bstrPropName, ref int plPropNum)
        {	//Здесь 1С ищет числовой идентификатор свойства по его текстовому имени
            InitByNeed();
            PyFindProp(bstrPropName, ref plPropNum);
        }

        public void GetPropName(int lPropNum, int lPropAlias, [MarshalAs(UnmanagedType.BStr)] ref string pbstrPropName)
        {	//Здесь 1С (теоретически) узнает имя свойства по его идентификатору. lPropAlias - номер псевдонима
            InitByNeed();
            PyGetPropName(lPropNum, lPropAlias, ref pbstrPropName);
        }

        public void GetPropVal(int lPropNum, [MarshalAs(UnmanagedType.Struct)] ref object pvarPropVal)
        {	//Здесь 1С узнает значения свойств
            InitByNeed();
            PyGetPropVal(lPropNum, ref pvarPropVal);
        }

        public void SetPropVal(int lPropNum, [MarshalAs(UnmanagedType.Struct)] ref object varPropVal)
        {	//Здесь 1С изменяет значения свойств
            InitByNeed();
            PySetPropVal(lPropNum, ref varPropVal);
        }

        public void IsPropReadable(int lPropNum, ref bool pboolPropRead)
        {	//Здесь 1С узнает, какие свойства доступны для чтения
            InitByNeed();
            PyIsPropReadable(lPropNum, ref pboolPropRead); // Все свойства доступны для чтения
        }

        public void IsPropWritable(int lPropNum, ref bool pboolPropWrite)
        {	//Здесь 1С узнает, какие свойства доступны для записи
            InitByNeed();
            PyIsPropWritable(lPropNum, ref pboolPropWrite); // Все свойства доступны для записи
        }

        #endregion

        #region "Методы"

        enum Methods
        {	//Числовые идентификаторы методов (процедур или функций) нашей внешней компоненты
            LastMethod = 0,
        }

        public void GetNMethods(ref int plMethods)
        {	//Здесь 1С получает количество доступных из ВК методов
            InitByNeed();
            PyGetNMethods(ref plMethods);
        }

        public void FindMethod([MarshalAs(UnmanagedType.BStr)] string bstrMethodName, ref int plMethodNum)
        {	//Здесь 1С получает числовой идентификатор метода (процедуры или функции) по имени (названию) процедуры или функции
            InitByNeed();
            PyFindMethod(bstrMethodName, ref plMethodNum);
        }

        public void GetMethodName(int lMethodNum, int lMethodAlias, [MarshalAs(UnmanagedType.BStr)] ref string pbstrMethodName)
        {	//Здесь 1С (теоретически) получает имя метода по его идентификатору. lMethodAlias - номер синонима.
            InitByNeed();
            PyGetMethodName(lMethodNum, lMethodAlias, ref pbstrMethodName);
        }

        public void GetNParams(int lMethodNum, ref int plParams)
        {	//Здесь 1С получает количество параметров у метода (процедуры или функции)
            InitByNeed();
            PyGetNParams(lMethodNum, ref plParams);
        }

        public void GetParamDefValue(int lMethodNum, int lParamNum, [MarshalAs(UnmanagedType.Struct)] ref object pvarParamDefValue)
        {	//Здесь 1С получает значения параметров процедуры или функции по умолчанию
            InitByNeed();
            PyGetParamDefValue(lMethodNum, lParamNum, ref pvarParamDefValue); //Нет значений по умолчанию
        }

        public void HasRetVal(int lMethodNum, ref bool pboolRetValue)
        {	//Здесь 1С узнает, возвращает ли метод значение (т.е. является процедурой или функцией)
            InitByNeed();
            PyHasRetVal(lMethodNum, ref pboolRetValue);  //Все методы у нас будут функциями (т.е. будут возвращать значение). 
        }

        public void CallAsProc(int lMethodNum, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {	//Здесь внешняя компонента выполняет код процедур. А процедур у нас нет.
            InitByNeed();
            PyCallAsProc(lMethodNum, ref pParams);
        }

        public void CallAsFunc(int lMethodNum, [MarshalAs(UnmanagedType.Struct)] ref object pvarRetValue, [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] pParams)
        {	//Здесь внешняя компонента выполняет код функций.
            InitByNeed();
            PyCallAsFunc(lMethodNum, ref pvarRetValue, ref pParams); //Возвращаемое значение метода для 1С			
        }

        #endregion

    }
}
