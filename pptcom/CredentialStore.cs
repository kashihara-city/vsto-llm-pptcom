using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Pptcom
{
    // Minimal wrapper over the Win32 credential API (generic credentials of the current user).
    internal static class CredentialStore
    {
        private const int CRED_TYPE_GENERIC = 1;
        private const int CRED_PERSIST_LOCAL_MACHINE = 2; // this user on this PC (does not roam)

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct CREDENTIAL
        {
            public int Flags;
            public int Type;
            public string TargetName;
            public string Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public int CredentialBlobSize;
            public IntPtr CredentialBlob;
            public int Persist;
            public int AttributeCount;
            public IntPtr Attributes;
            public string TargetAlias;
            public string UserName;
        }

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credential);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredWrite(ref CREDENTIAL credential, int flags);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredDelete(string target, int type, int flags);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("advapi32.dll")]
        private static extern void CredFree(IntPtr buffer);

        public static string Read(string target)
        {
            if (!CredRead(target, CRED_TYPE_GENERIC, 0, out IntPtr p))
            {
                return null;
            }

            try
            {
                CREDENTIAL c = Marshal.PtrToStructure<CREDENTIAL>(p);
                if (c.CredentialBlob == IntPtr.Zero || c.CredentialBlobSize == 0)
                {
                    return "";
                }

                byte[] bytes = new byte[c.CredentialBlobSize];
                Marshal.Copy(c.CredentialBlob, bytes, 0, bytes.Length);
                return Encoding.Unicode.GetString(bytes);
            }
            finally
            {
                CredFree(p);
            }
        }

        public static bool Write(string target, string userName, string secret)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(secret);
            IntPtr blob = Marshal.AllocHGlobal(bytes.Length);
            try
            {
                Marshal.Copy(bytes, 0, blob, bytes.Length);
                CREDENTIAL c = new()
                {
                    Type = CRED_TYPE_GENERIC,
                    TargetName = target,
                    UserName = userName,
                    Comment = "pptcom (vsto-llm-pptcom): API key for the LLM endpoint",
                    CredentialBlob = blob,
                    CredentialBlobSize = bytes.Length,
                    Persist = CRED_PERSIST_LOCAL_MACHINE,
                };
                return CredWrite(ref c, 0);
            }
            finally
            {
                Marshal.FreeHGlobal(blob);
            }
        }

        public static bool Delete(string target)
        {
            return CredDelete(target, CRED_TYPE_GENERIC, 0) || Read(target) is null;
        }
    }
}
