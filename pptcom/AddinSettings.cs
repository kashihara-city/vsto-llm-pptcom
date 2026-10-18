using System;
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

        public bool HasEndpoint => Endpoint.Length > 0;
        public bool HasApiKey => ApiKey.Length > 0;

        public static AddinSettings Load()
        {
            AddinSettings s = new();
            try
            {
                using RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryPath);
                if (key is not null)
                {
                    s.Endpoint = ((key.GetValue("endpoint") as string) ?? "").Trim().TrimEnd('/');
                    s.ChatModel = ((key.GetValue("chat_model") as string) ?? "").Trim();
                    s.ImageModel = ((key.GetValue("image_model") as string) ?? "").Trim();
                    s.ExternalControl = key.GetValue("external_control") is int ec && ec != 0;
                    s.SaveProgress = key.GetValue("save_progress") is int sp && sp != 0;
                }
            }
            catch (Exception ex) when (IsRegistryFailure(ex))
            {
                // unreadable: start empty
            }
            s.ApiKey = CredentialStore.Read(CredentialTarget) ?? "";
            return s;
        }

        // Endpoint and models (the API key is saved separately with SaveApiKey).
        public void Save()
        {
            try
            {
                using RegistryKey key = Registry.CurrentUser.CreateSubKey(RegistryPath);
                key.SetValue("endpoint", Endpoint, RegistryValueKind.String);
                key.SetValue("chat_model", ChatModel, RegistryValueKind.String);
                key.SetValue("image_model", ImageModel, RegistryValueKind.String);
                key.SetValue("external_control", ExternalControl ? 1 : 0, RegistryValueKind.DWord);
                key.SetValue("save_progress", SaveProgress ? 1 : 0, RegistryValueKind.DWord);
            }
            catch (Exception ex) when (IsRegistryFailure(ex))
            {
            }
        }

        // What Registry throws when the key cannot be read or written (a policy, a locked hive).
        private static bool IsRegistryFailure(Exception ex)
        {
            return ex is System.Security.SecurityException or UnauthorizedAccessException or System.IO.IOException;
        }

        // Empty = remove the stored key.
        public bool SaveApiKey(string apiKey)
        {
            apiKey = apiKey.Trim();
            bool ok = apiKey.Length == 0 ? CredentialStore.Delete(CredentialTarget) : CredentialStore.Write(CredentialTarget, "api_key", apiKey);
            if (ok)
            {
                ApiKey = apiKey;
            }

            return ok;
        }
    }
}
