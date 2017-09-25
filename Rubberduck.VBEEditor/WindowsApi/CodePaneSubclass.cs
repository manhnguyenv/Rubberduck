using System;
using System.Windows.Forms;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.WindowsApi
{
    //Stub for code pane replacement.  :-)
    internal class CodePaneSubclass : FocusSource
    {
        private readonly ICodePane _pane;
        private readonly IntPtr _hwnd;
        private readonly UserControl _control;

        public ICodePane CodePane => _pane;

        internal CodePaneSubclass(IntPtr hwnd, ICodePane pane, UserControl control) 
            : base(hwnd)
        {
            _hwnd = hwnd;
            _pane = pane;
            _control = control;
            User32.SetParent(control.Handle, hwnd);

            var rect = new User32.RECT();
            User32.GetClientRect(hwnd, ref rect);
            User32.MoveWindow(control.Handle, 0, 0, rect.Right - rect.Left, rect.Bottom - rect.Top, true);
            //User32.SetWindowPos(_control.Handle, new IntPtr(-1), 0, 0, rect.Right - rect.Left, rect.Bottom - rect.Top, 0);
        }

        protected override void HandleResized(int width, int height)
        {
            User32.MoveWindow(_control.Handle, 0, 0, width, height, true);
        }

        protected override void DispatchFocusEvent(FocusType type)
        {
            var window = VBENativeServices.GetWindowInfoFromHwnd(Hwnd);
            if (!window.HasValue)
            {
                return;
            }
            OnFocusChange(new WindowChangedEventArgs(window.Value.Hwnd, window.Value.Window, _pane, type));
        }
    }
}
