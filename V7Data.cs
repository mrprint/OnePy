using System;
using Enterprise.AddIn;

namespace OnePy
{
	internal class V7Data
	{

		public static object V7Object
		{
			get
			{
				return m_V7Object;
			}
			set
			{
				m_V7Object = value;
				// Вызываем неявно QueryInterface
				m_ErrorInfo = (IErrorLog)value;
				m_AsyncEvent = (IAsyncEvent)value;
				m_StatusLine = (IStatusLine)value;
                m_ExtWndsSupport = (IExtWndsSupport)value;
                m_PropertyProfile = (IPropertyProfile)value;
			}
		}

        //public static object obj1C
        //{
        //    get
        //    {
        //        return m_obj1C;
        //    }
        //    set 
        //    {
        //        m_obj1C = value;
        //    }
        //}

		public static IErrorLog ErrorLog
		{
			get
			{
				return m_ErrorInfo;
			}
		}

		public static IAsyncEvent AsyncEvent
		{
			get
			{
				return m_AsyncEvent;
			}
		}

		public static IStatusLine StatusLine
		{
			get
			{
				return m_StatusLine;
			}
		}

        public static IExtWndsSupport ExtWndsSupport
        {
            get
            {
                return m_ExtWndsSupport;
            }
        }

        public static IPropertyProfile PropertyProfile
        {
            get
            {
                return m_PropertyProfile;
            }
        }

        public static void Clean()
        {
            m_AsyncEvent = null;
            m_ErrorInfo = null;
            m_ExtWndsSupport = null;
            m_PropertyProfile = null;
            m_StatusLine = null;
            m_V7Object = null;
        }


		private static object m_V7Object;
        //private static object m_obj1C; // устанавливается в конструкторе InteractData
		private static IErrorLog m_ErrorInfo;
		private static IAsyncEvent m_AsyncEvent;
		private static IStatusLine m_StatusLine;
        private static IExtWndsSupport m_ExtWndsSupport;
        private static IPropertyProfile m_PropertyProfile;
	}
}
