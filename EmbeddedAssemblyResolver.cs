using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ВыполнитьЗадачиSolidWorks
{
    internal static class EmbeddedAssemblyResolver
    {
        private const string ResourcePrefix = "EmbeddedAssembly.";
        private static readonly object SyncRoot = new object();
        private static readonly Dictionary<string, Assembly> Loaded = new Dictionary<string, Assembly>(StringComparer.OrdinalIgnoreCase);
        private static bool initialized;

        public static void Initialize()
        {
            if (initialized) return;

            lock (SyncRoot)
            {
                if (initialized) return;

                AppDomain.CurrentDomain.AssemblyResolve += Resolve;
                initialized = true;
            }
        }

        private static Assembly Resolve(object sender, ResolveEventArgs args)
        {
            string dllName = new AssemblyName(args.Name).Name + ".dll";

            lock (SyncRoot)
            {
                if (Loaded.TryGetValue(dllName, out Assembly cached))
                {
                    return cached;
                }

                Assembly current = Assembly.GetExecutingAssembly();
                string resourceName = ResourcePrefix + dllName;

                using (Stream stream = current.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        return null;
                    }

                    byte[] bytes = new byte[stream.Length];
                    int offset = 0;

                    while (offset < bytes.Length)
                    {
                        int read = stream.Read(bytes, offset, bytes.Length - offset);
                        if (read == 0) break;
                        offset += read;
                    }

                    Assembly assembly = Assembly.Load(bytes);
                    Loaded[dllName] = assembly;
                    return assembly;
                }
            }
        }
    }
}
