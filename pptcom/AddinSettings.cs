using System;
using Microsoft.Win32;

namespace Pptcom
{
    // Connection settings. Nothing environment-specific is built in; the add-in, the ComScriptRunner
    // rig and scripts/test/tool_loop.py all read the same places:
    //
    //   HKCU\Software\KashiharaCity\pptcom  endpoint, chat_model, image_model (REG_SZ; image_model empty: no images);
    //                                        image_prompt (REG_SZ, one of ImagePrompts, add-in only);
    //                                        external_control, save_progress, open_pane (REG_DWORD 0/1, add-in only);
    //                                        context_limit (REG_DWORD, tokens, add-in only);
    //                                        AppVersion (REG_SZ, the installed version, written by the MSI) and
    //                                        LatestVersion (REG_SZ, the version an administrator hands out by group
    //                                        policy), for the logon script that installs it; the add-in reads neither
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
        // ImagePrompt: which rules and sizes imagegen gets (Agent\image-prompts\<name>.json, see AgentTools), as each
        // image model wants prompts of its own. One of ImagePrompts, which the page's drop-down lists too (index.html);
        // any other value is read as the first.
        public string ImagePrompt = ImagePrompts[0];
        public static readonly string[] ImagePrompts = ["zimage"];

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
        // OpenPane: each presentation window opens with its task pane shown (on unless the user turns it off: the
        // pane always opened before this setting). Off, the pane waits hidden for the ribbon's AIアシスタント button.
        public bool OpenPane = true;
        // ContextLimit: when a turn's prompt passes this many tokens, the request's conversation is shortened after it
        // (ThisAddIn.CompactAsync). Set to three quarters of what the chat model takes in when the model is chosen
        // (ThisAddIn.SetContextLimitFromModel), DefaultContextLimit when the endpoint does not say; the user may change it.
        // ContextLimitSaved: the registry has one (set once from the model otherwise).
        public int ContextLimit = DefaultContextLimit;
        public bool ContextLimitSaved;
        public const int DefaultContextLimit = 100000;
        public const int MinContextLimit = 8000;
        public const int MaxContextLimit = 10000000;

        // %LOCALAPPDATA%\KashiharaCity\pptcom: app\ (installed by the MSI), temp\ (temporary files and the
        // WebView2 profile; safe to delete while PowerPoint is closed) and, with SaveProgress, the records.
        public static readonly string DataFolder = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "KashiharaCity", "pptcom");

        public bool HasEndpoint => Endpoint.Length > 0;
        public bool HasApiKey => ApiKey.Length > 0;

        // The API key goes only over https (OpenAiCompatibleClient refuses the rest, localhost too): while one is set,
        // an http endpoint is not used, and none is set for an http endpoint.
        public static bool IsHttp(string endpoint)
        {
            return endpoint.StartsWith("http://", StringComparison.OrdinalIgnoreCase);
        }

        public bool KeyOverHttp => HasApiKey && IsHttp(Endpoint);

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
                    string imagePrompt = ((key.GetValue("image_prompt") as string) ?? "").Trim();
                    s.ImagePrompt = Array.IndexOf(ImagePrompts, imagePrompt) >= 0 ? imagePrompt : ImagePrompts[0];
                    s.ExternalControl = key.GetValue("external_control") is int ec && ec != 0;
                    s.SaveProgress = key.GetValue("save_progress") is int sp && sp != 0;
                    s.OpenPane = key.GetValue("open_pane") is not int op || op != 0;
                    if (key.GetValue("context_limit") is int cl && cl >= MinContextLimit && cl <= MaxContextLimit)
                    {
                        s.ContextLimit = cl;
                        s.ContextLimitSaved = true;
                    }
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
                key.SetValue("image_prompt", ImagePrompt, RegistryValueKind.String);
                key.SetValue("external_control", ExternalControl ? 1 : 0, RegistryValueKind.DWord);
                key.SetValue("save_progress", SaveProgress ? 1 : 0, RegistryValueKind.DWord);
                key.SetValue("open_pane", OpenPane ? 1 : 0, RegistryValueKind.DWord);
                if (ContextLimitSaved)
                {
                    key.SetValue("context_limit", ContextLimit, RegistryValueKind.DWord);
                }
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
