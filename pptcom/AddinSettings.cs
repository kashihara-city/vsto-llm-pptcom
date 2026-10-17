using System;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

namespace Pptcom
{
    // Connection settings. Nothing environment-specific is built in; the add-in, the ComScriptRunner
    // rig and scripts/test/tool_loop.py all read the same places:
    //
    //   HKCU\Software\KashiharaCity\pptcom  endpoint, chat_model, image_model (REG_SZ);
    //                                        external_control, save_progress (REG_DWORD 0/1, add-in only)
    //   Windows Credential Manager           generic credential "KashiharaCity/pptcom:api_key" (the API key,
    //                                        sent as the Bearer token; never written to disk in plain text)
    //
    // The endpoint and both models are chosen in the task pane; the API key is entered there too
    // (masked, never shown back) or in Control Panel > Credential Manager > Windows Credentials.
    internal sealed class AddinSettings
    {
        public string Endpoint = "";
        public string ApiKey = "";
        public string ChatModel = "";
        public string ImageModel = "";

        public const string RegistryPath = @"Software\KashiharaCity\pptcom";
        public const string CredentialTarget = "KashiharaCity/pptcom:api_key";

        // Where to tell the user to look (messages in the task pane / runner).
        public const string Location = @"HKCU\" + RegistryPath + " / Windows Credential Manager \"" + CredentialTarget + "\"";

        // Both off unless the user turns them on in the task pane's settings.
        // ExternalControl: other programs may drive the add-in through its COM automation object
        // (AddinAutomation) - for tests that run the real add-in from a script.
        public bool ExternalControl;
        // SaveProgress: keep session.log, com_usage.jsonl and a folder per request (generated-pptx)
        // under DataFolder. They hold the requests, the model's reasoning and the deck's text.
        public bool SaveProgress;

        // %LOCALAPPDATA%\KashiharaCity\pptcom: app\ (installed by the MSI), temp\ (temporary files and the
        // WebView2 profile; safe to delete while PowerPoint is closed) and, with SaveProgress, the records.
        public static readonly string DataFolder = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KashiharaCity", "pptcom");

        public bool HasEndpoint { get { return Endpoint.Length > 0; } }
        public bool HasApiKey { get { return ApiKey.Length > 0; } }

        public static AddinSettings Load()
        {
            var s = new AddinSettings();
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(RegistryPath))
                {
                    if (key != null)
                    {
                        s.Endpoint = ((key.GetValue("endpoint") as string) ?? "").Trim().TrimEnd('/');
                        s.ChatModel = ((key.GetValue("chat_model") as string) ?? "").Trim();
                        s.ImageModel = ((key.GetValue("image_model") as string) ?? "").Trim();
                        s.ExternalControl = key.GetValue("external_control") is int ec && ec != 0;
                        s.SaveProgress = key.GetValue("save_progress") is int sp && sp != 0;
                    }
                }
            }
            catch { /* unreadable: start empty */ }
            s.ApiKey = CredentialStore.Read(CredentialTarget) ?? "";
            return s;
        }

        // Endpoint and models (the API key is saved separately with SaveApiKey).
        public void Save()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(RegistryPath))
                {
                    key.SetValue("endpoint", Endpoint, RegistryValueKind.String);
                    key.SetValue("chat_model", ChatModel, RegistryValueKind.String);
                    key.SetValue("image_model", ImageModel, RegistryValueKind.String);
                    key.SetValue("external_control", ExternalControl ? 1 : 0, RegistryValueKind.DWord);
                    key.SetValue("save_progress", SaveProgress ? 1 : 0, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        // Empty = remove the stored key.
        public bool SaveApiKey(string apiKey)
        {
            apiKey = (apiKey ?? "").Trim();
            bool ok = apiKey.Length == 0 ? CredentialStore.Delete(CredentialTarget) : CredentialStore.Write(CredentialTarget, "api_key", apiKey);
            if (ok) ApiKey = apiKey;
            return ok;
        }
    }

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

        [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credential);

        [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredWrite(ref CREDENTIAL credential, int flags);

        [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool CredDelete(string target, int type, int flags);

        [DllImport("advapi32.dll")]
        private static extern void CredFree(IntPtr buffer);

        public static string Read(string target)
        {
            IntPtr p;
            if (!CredRead(target, CRED_TYPE_GENERIC, 0, out p)) return null;
            try
            {
                var c = (CREDENTIAL)Marshal.PtrToStructure(p, typeof(CREDENTIAL));
                if (c.CredentialBlob == IntPtr.Zero || c.CredentialBlobSize == 0) return "";
                var bytes = new byte[c.CredentialBlobSize];
                Marshal.Copy(c.CredentialBlob, bytes, 0, bytes.Length);
                return Encoding.Unicode.GetString(bytes);
            }
            finally { CredFree(p); }
        }

        public static bool Write(string target, string userName, string secret)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(secret);
            IntPtr blob = Marshal.AllocHGlobal(bytes.Length);
            try
            {
                Marshal.Copy(bytes, 0, blob, bytes.Length);
                var c = new CREDENTIAL
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
            finally { Marshal.FreeHGlobal(blob); }
        }

        public static bool Delete(string target)
        {
            return CredDelete(target, CRED_TYPE_GENERIC, 0) || Read(target) == null;
        }
    }
}
