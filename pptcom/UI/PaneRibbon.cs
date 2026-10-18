using System;
using System.Runtime.InteropServices;
using System.Security;
using Office = Microsoft.Office.Core;

namespace Pptcom.UI
{
    // The add-in's button on the ribbon's ホーム tab: shows or hides the task pane of the window in front (each
    // window has its own, see ThisAddIn.PaneFor). Its pressed state follows that pane; ThisAddIn refreshes it when
    // another window comes to the front and when a pane is shown or closed. Office calls these members through COM
    // (IDispatch), so the class is public and visible to COM, as AddinAutomation is. The labels are written into the
    // ribbon's XML in the add-in's language when Office asks for it.
    [ComVisible(true)]
    public sealed class PaneRibbon : Office.IRibbonExtensibility
    {
        private const string ButtonId = "pptcomPane";

        private readonly Func<bool> _isShown;
        private readonly Action<bool> _show;
        private Office.IRibbonUI? _ui;

        internal PaneRibbon(Func<bool> isShown, Action<bool> show)
        {
            _isShown = isShown;
            _show = show;
        }

        public string GetCustomUI(string RibbonID)
        {
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnLoad\">" +
                   "<ribbon><tabs><tab idMso=\"TabHome\">" +
                   "<group id=\"pptcomGroup\" label=\"" + Attribute(Texts.Get("RIBBON_GROUP")) + "\">" +
                   "<toggleButton id=\"" + ButtonId + "\" size=\"large\" imageMso=\"ReviewNewComment\" label=\"" + Attribute(Texts.Get("RIBBON_PANE")) + "\" " +
                   "supertip=\"" + Attribute(Texts.Get("RIBBON_PANE_TIP")) + "\" getPressed=\"GetPressed\" onAction=\"OnTogglePane\"/>" +
                   "</group></tab></tabs></ribbon></customUI>";
        }

        private static string Attribute(string text)
        {
            return SecurityElement.Escape(text);
        }

        public void OnLoad(Office.IRibbonUI ribbon)
        {
            _ui = ribbon;
        }

        public bool GetPressed(Office.IRibbonControl control)
        {
            return IsButton(control) && _isShown();
        }

        public void OnTogglePane(Office.IRibbonControl control, bool pressed)
        {
            if (IsButton(control))
            {
                _show(pressed);
            }
        }

        private static bool IsButton(Office.IRibbonControl control)
        {
            return control?.Id == ButtonId;
        }

        // The button shows the state of the pane in front again.
        internal void Refresh()
        {
            _ui?.InvalidateControl(ButtonId);
        }
    }
}
