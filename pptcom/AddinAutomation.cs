using System;
using System.Runtime.InteropServices;

namespace Pptcom
{
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
            {
                throw new InvalidOperationException("外部からの操作は無効です。作業ウィンドウの接続設定で「外部からの操作を許可」をオンにしてください。");
            }
        }

        public void Send(string request)
        {
            RequireEnabled();
            if (string.IsNullOrWhiteSpace(request))
            {
                throw new ArgumentException("request is empty", nameof(request));
            }

            if (_addin.IsBusy)
            {
                throw new InvalidOperationException("実行中の依頼があります。IsBusy が False になってから送ってください。");
            }

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
            if (_addin.IsBusy)
            {
                throw new InvalidOperationException("実行中は新しい会話を始められません。");
            }

            _addin.NewConversationFromOutside();
        }

        public bool IsBusy
        {
            get
            {
                RequireEnabled();
                return _addin.IsBusy;
            }
        }

        public void Attach(string path)
        {
            RequireEnabled();
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("path is empty", nameof(path));
            }

            if (_addin.IsBusy)
            {
                throw new InvalidOperationException("実行中は資料を添付できません。");
            }

            _addin.AttachFromOutside(path.Trim());
        }

        public string Documents
        {
            get
            {
                RequireEnabled();
                return _addin.DocumentsState().ToString(Newtonsoft.Json.Formatting.None);
            }
        }

        public string LastRun
        {
            get
            {
                RequireEnabled();
                Newtonsoft.Json.Linq.JObject? summary = _addin.LastRunSummary;
                return summary is not null ? summary.ToString(Newtonsoft.Json.Formatting.None) : "";
            }
        }
    }
}
