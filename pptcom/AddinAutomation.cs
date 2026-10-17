using System;
using System.Runtime.InteropServices;

namespace Pptcom
{
    // The add-in's COM automation object, for driving the real add-in from a script (tests that need
    // the add-in as the user runs it: the conversation, the task pane, the same PowerPoint process).
    // A script reaches it through PowerPoint:
    //
    //   $ppt = New-Object -ComObject PowerPoint.Application
    //   $ai  = $ppt.COMAddIns.Item("pptcom").Object
    //   $ai.Send("バナナについて3枚にまとめて"); while ($ai.IsBusy) { Start-Sleep 1 }; $ai.LastRun
    //
    // Every call is refused unless the user turned on 外部からの操作を許可 in the task pane
    // (AddinSettings.ExternalControl, off by default): with it on, any program running as this user
    // could send requests with the stored API key and change the open decks. A request sent this way
    // shows in the task pane like a typed one, marked as coming from another program.
    [ComVisible(true)]
    [Guid("5D551056-7DCE-47CA-9023-BA48697065BC")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IPptAiAutomation
    {
        // Same as typing the request and pressing send. Returns at once; the request runs in PowerPoint.
        void Send(string request);
        // Same as the stop button.
        void Cancel();
        // Same as 新しい会話.
        void NewConversation();
        // True while a request runs.
        bool IsBusy { get; }
        // The last finished request as JSON (summary.json's fields, the tool lines the pane showed and the
        // model's reply; run_folder only with 途中経過を保存 on). Empty before the first request.
        string LastRun { get; }
        // Same as dropping the file on the task pane. Returns at once: the pane converts it, and it shows
        // up in Documents (or under "failed" there) when done.
        void Attach(string path);
        // JSON: {documents: [{id, name, format, lines, images}], converting: [names], failed: ["name: reason"]}.
        string Documents { get; }
    }

    // StandardOleMarshalObject: calls from another process are run on PowerPoint's UI thread, where
    // the add-in and the task pane live.
    [ComVisible(true)]
    [Guid("02B79EA6-84F9-4596-969C-EC55B0591F47")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class AddinAutomation : StandardOleMarshalObject, IPptAiAutomation
    {
        private readonly ThisAddIn _addin;

        internal AddinAutomation(ThisAddIn addin)
        {
            _addin = addin;
        }

        private void RequireEnabled()
        {
            if (!_addin.ExternalControlEnabled)
                throw new InvalidOperationException("外部からの操作は無効です。作業ウィンドウの接続設定で「外部からの操作を許可」をオンにしてください。");
        }

        public void Send(string request)
        {
            RequireEnabled();
            if (string.IsNullOrWhiteSpace(request)) throw new ArgumentException("request is empty");
            if (_addin.IsBusy) throw new InvalidOperationException("実行中の依頼があります。IsBusy が False になってから送ってください。");
            _addin.SendFromOutside(request.Trim());
        }

        public void Cancel()
        {
            RequireEnabled();
            _addin.CancelFromOutside();
        }

        public void NewConversation()
        {
            RequireEnabled();
            if (_addin.IsBusy) throw new InvalidOperationException("実行中は新しい会話を始められません。");
            _addin.NewConversationFromOutside();
        }

        public bool IsBusy
        {
            get { RequireEnabled(); return _addin.IsBusy; }
        }

        public void Attach(string path)
        {
            RequireEnabled();
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("path is empty");
            if (_addin.IsBusy) throw new InvalidOperationException("実行中は資料を添付できません。");
            _addin.AttachFromOutside(path.Trim());
        }

        public string Documents
        {
            get { RequireEnabled(); return _addin.DocumentsState().ToString(Newtonsoft.Json.Formatting.None); }
        }

        public string LastRun
        {
            get
            {
                RequireEnabled();
                var summary = _addin.LastRunSummary;
                return summary != null ? summary.ToString(Newtonsoft.Json.Formatting.None) : "";
            }
        }
    }
}
