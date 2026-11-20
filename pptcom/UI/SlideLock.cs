using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Pptcom.UI
{
    // Keeps the user's hands off a presentation's slides while a request works on it: what the model saw of them and
    // the shapes its variables hold stay as they are until it is done. PowerPoint has nothing like a worksheet's
    // protection, so the part of each of the presentation's windows that holds the slides - the document window
    // (class mdiClass, which draws the slide, the thumbnails and the notes itself) - is disabled. Windows then sends it
    // no mouse input, takes the keyboard focus from it and does not let it come back, while the object model goes on
    // working on it. The rest of the frame around it stays as it is: the task pane (its stop button), and also the
    // ribbon and its shortcuts, which this does not hold back.
    // Used on PowerPoint's thread, where its windows live: Windows takes the focus from a window being disabled only
    // for the thread that disables it.
    internal sealed class SlideLock(string documentClass = SlideLock.DocumentWindowClass)
    {
        public const string DocumentWindowClass = "mdiClass";

        // the window a presentation's window is known by -> its document window, disabled here
        private readonly Dictionary<IntPtr, IntPtr> _held = [];
        // the document window that had the keyboard focus when it was disabled; it gets it back (Release)
        private IntPtr _hadFocus;

        public int Count => _held.Count;

        public bool Holds(IntPtr window)
        {
            return _held.ContainsKey(window);
        }

        // Disables the document window of each of these windows that is not held yet; returns how many it disabled now.
        // A window is the document window itself (PowerPoint's DocumentWindow.HWND is) or one that has it inside (its
        // frame). One without a document window, or whose document window something else has disabled, is left alone
        // and not held.
        public int Lock(IEnumerable<IntPtr> windows)
        {
            int added = 0;
            foreach (IntPtr window in windows)
            {
                if (_held.ContainsKey(window))
                {
                    continue;
                }

                IntPtr document = Find(window, depth: 0);
                if (document == IntPtr.Zero || !IsWindowEnabled(document))
                {
                    continue;
                }

                bool focused = Within(document, GetFocus());
                _ = EnableWindow(document, enable: false);
                if (Within(document, GetFocus()))
                {
                    _ = SetFocus(IntPtr.Zero); // Windows did not take it away: no window has it now
                }

                if (focused)
                {
                    _hadFocus = document;
                }

                _held[window] = document;
                added++;
            }

            return added;
        }

        // Enables what Lock disabled. The window that had the keyboard focus gets it back, unless the user has moved it
        // meanwhile: to another window of this thread, or into the task pane beside it.
        public void Release()
        {
            foreach (IntPtr document in _held.Values)
            {
                if (IsWindow(document))
                {
                    _ = EnableWindow(document, enable: true);
                }
            }

            _held.Clear();
            if (_hadFocus != IntPtr.Zero && IsWindow(_hadFocus) && GetFocus() == IntPtr.Zero && !TypingBeside(_hadFocus))
            {
                _ = SetFocus(_hadFocus);
            }

            _hadFocus = IntPtr.Zero;
        }

        // The window itself or the first window of the document class below it (a frame holds one).
        private IntPtr Find(IntPtr window, int depth)
        {
            if (window == IntPtr.Zero || !IsWindow(window))
            {
                return IntPtr.Zero;
            }

            if (string.Equals(ClassOf(window), documentClass, StringComparison.OrdinalIgnoreCase))
            {
                return window;
            }

            if (depth >= MaxDepth)
            {
                return IntPtr.Zero;
            }

            for (IntPtr child = GetWindow(window, GwChild); child != IntPtr.Zero; child = GetWindow(child, GwHwndNext))
            {
                IntPtr found = Find(child, depth + 1);
                if (found != IntPtr.Zero)
                {
                    return found;
                }
            }

            return IntPtr.Zero;
        }

        private static bool Within(IntPtr document, IntPtr window)
        {
            return window != IntPtr.Zero && (window == document || IsChild(document, window));
        }

        // The keyboard focus is in a window of another thread in the same frame as the document window: the page of
        // the task pane, which runs in the browser's process (this thread's own GetFocus says nothing of it).
        private static bool TypingBeside(IntPtr document)
        {
            GuiThreadInfo info = new() { Size = Marshal.SizeOf<GuiThreadInfo>() };
            return GetGUIThreadInfo(0, ref info) && info.Focus != IntPtr.Zero && IsChild(GetAncestor(document, GaRoot), info.Focus);
        }

        private static string ClassOf(IntPtr window)
        {
            char[] name = new char[128];
            int length = GetClassName(window, name, name.Length);
            return length > 0 ? new string(name, 0, length) : "";
        }

        // PowerPoint's document window is two levels below its frame (MDIClient, then mdiClass).
        private const int MaxDepth = 4;
        private const uint GwHwndNext = 2, GwChild = 5, GaRoot = 2;

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left, Top, Right, Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct GuiThreadInfo
        {
            public int Size;
            public int Flags;
            public IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
            public Rect CaretRect;
        }

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnableWindow(IntPtr window, [MarshalAs(UnmanagedType.Bool)] bool enable);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowEnabled(IntPtr window);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr window);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsChild(IntPtr parent, IntPtr window);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr window, uint command);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr window, uint flags);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr window);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetGUIThreadInfo(uint thread, ref GuiThreadInfo info);

        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, [Out] char[] name, int capacity);
    }
}
