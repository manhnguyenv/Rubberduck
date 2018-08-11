﻿using System;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.WindowsApi
{
    //Stub for code pane replacement.  :-)
    internal class CodePaneSubclass : VbeAttachableSubclass<ICodePane>, IWindowEventProvider
    {       
        public event EventHandler CaptionChanged;
        public event EventHandler<KeyPressEventArgs> KeyDown;
 
        internal CodePaneSubclass(IntPtr hwnd, ICodePane pane) : base(hwnd)
        {
            VbeObject = pane;
        }

        protected void OnKeyDown(KeyPressEventArgs eventArgs)
        {
            KeyDown?.Invoke(this, eventArgs);
        }

        public override int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            KeyPressEventArgs args;
            switch ((WM)msg)
            {
                case WM.CHAR:
                    args = new KeyPressEventArgs(hWnd, wParam, lParam, (char)wParam);
                    OnKeyDown(args);
                    if (args.Handled) { return 0; }
                    break;
                case WM.KEYDOWN:
                    args = new KeyPressEventArgs(hWnd, wParam, lParam);
                    OnKeyDown(args);
                    if (args.Handled) { return 0; }
                    break;
                case WM.SETTEXT:
                    if (!HasValidVbeObject)
                    {
                        CaptionChanged?.Invoke(this, null);
                    }
                    break;
            }
            return base.SubClassProc(hWnd, msg, wParam, lParam, uIdSubclass, dwRefData);
        }

        private bool _disposed;
        protected override void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                CaptionChanged = delegate { };
                KeyDown = delegate { };
            }

            base.Dispose(disposing);
            _disposed = true;
        }

        protected override void DispatchFocusEvent(FocusType type)
        {
            OnFocusChange(new WindowChangedEventArgs(Hwnd, type));
        }
    }
}
