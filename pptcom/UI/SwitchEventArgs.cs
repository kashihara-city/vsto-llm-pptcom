using System;

namespace Pptcom.UI
{
    // A switch of the settings, turned on or off.
    internal sealed class SwitchEventArgs : EventArgs
    {
        public SwitchEventArgs(bool on)
        {
            On = on;
        }
        public bool On { get; }
    }
}
