using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Reflection;
using System.Collections;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Scripting;
using Enterprise.AddIn;

namespace OnePy
{
    class MessageException : Exception
    {
    }

    class AssemblyStreamContentProvider : StreamContentProvider
    {
        private readonly Assembly assembly;
        private readonly string filename;

        public AssemblyStreamContentProvider(string fileName)
        {
            filename = fileName;
            assembly = Assembly.GetCallingAssembly();
        }

#if (ONEPY45)
        public override Stream GetStream()
        {
            MemoryStream uncompressed = new MemoryStream();
            using (Stream ms = assembly.GetManifestResourceStream(filename))
            {
                using (GZipStream stream = new GZipStream(ms, CompressionMode.Decompress))
                {
                    stream.CopyTo(uncompressed);
                }
            }
            uncompressed.Seek(0, SeekOrigin.Begin);
            return uncompressed;
        }
#else
#if (ONEPY35)
        public override Stream GetStream()
        {
            MemoryStream uncompressed = new MemoryStream();
            using (Stream stream = assembly.GetManifestResourceStream(filename))
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    byte[] buffer = new byte[4096];
                    int read = 0;

                    while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        ms.Write(buffer, 0, read);
                    }
                    ms.Seek(0, SeekOrigin.Begin);
                    using (GZipStream gzipStream = new GZipStream(ms, CompressionMode.Decompress, false))
                    {
                        read = 0;
                        buffer = new byte[4096];

                        while ((read = gzipStream.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            uncompressed.Write(buffer, 0, read);
                        }
                    }
                }
            }
            uncompressed.Seek(0, SeekOrigin.Begin);
            return uncompressed;
        }
#endif
#endif

    }

    static class Log
    {
        private static object sync = new object();
        private static string dir = "";

        public static void PrepareDir()
        {
            string onepydir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string logspath = Path.Combine(onepydir, "OnePy");
            if (!Directory.Exists(logspath))
                Directory.CreateDirectory(logspath);
            dir = logspath;
        }

        public static void Write(string msg)
        {
            try
            {

                string filename = Path.Combine(dir, string.Format("dotnet.log", DateTime.Now));
                string fullText = string.Format("[{0:dd.MM.yyy HH:mm:ss.fff}] {1}\r\n", DateTime.Now, msg);
                lock (sync)
                {
                    File.AppendAllText(filename, fullText, Encoding.GetEncoding("Windows-1251"));
                }
            }
            catch
            {
            }
        }

        public static void DWrite(string msg)
        {
#if (DEBUG)
            Write(msg);
#endif
        }
    }


#if (ONEPY45)
    public partial class OnePy45
#else
#if (ONEPY35)
    public partial class OnePy35
#endif
#endif
    {
        static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        void ProcessError(Exception e)
        {
            if (e.Source != null)
            {
                Log.Write(e.Message + "\n" + e.StackTrace);
                if (Initialized)
                {
                    System.Runtime.InteropServices.ComTypes.EXCEPINFO ei = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                    ei.wCode = 1006; //Вид пиктограммы
                    ei.bstrDescription = e.Message; //.ToString()
                    ei.bstrSource = c_AddinName;
                    V7Data.ErrorLog.AddError("", ref ei);
                    //prClearExcInfo();
                    //throw new System.Exception("An exception has occurred.");
                }
            }
        }


    }
}