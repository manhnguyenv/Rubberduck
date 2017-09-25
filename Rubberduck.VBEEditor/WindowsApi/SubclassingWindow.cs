using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Rubberduck.VBEditor.WindowsApi
{
    public abstract class SubclassingWindow : IDisposable
    {
        /// <remarks>
        /// https://msdn.microsoft.com/en-us/library/windows/desktop/ms632612(v=vs.85).aspx
        /// </remarks>>
        private struct WindowPos
        {
            public IntPtr hwnd;
            public IntPtr hwndInsertAfter;
            public int x;
            public int y;
            public int cx;
            public int cy;
            public uint flags;
        }

        private readonly IntPtr _subclassId;
        private readonly SubClassCallback _wndProc;
        private bool _listening;
        private GCHandle _thisHandle;

        private readonly object _subclassLock = new object();

        protected delegate int SubClassCallback(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int SetWindowSubclass(IntPtr hWnd, SubClassCallback newProc, IntPtr uIdSubclass, IntPtr dwRefData);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int RemoveWindowSubclass(IntPtr hWnd, SubClassCallback newProc, IntPtr uIdSubclass);

        [DllImport("ComCtl32.dll", CharSet = CharSet.Auto)]
        private static extern int DefSubclassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam);

        protected IntPtr Hwnd { get; }

        protected SubclassingWindow(IntPtr subclassId, IntPtr hWnd)
        {
            _subclassId = subclassId;
            Hwnd = hWnd;
            _wndProc = SubClassProc;
            AssignHandle();
        }

        private bool _disposed;
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            ReleaseHandle();
            _thisHandle.Free();

            _disposed = true;
        }

        private void AssignHandle()
        {
            lock (_subclassLock)
            {
                var result = SetWindowSubclass(Hwnd, _wndProc, _subclassId, IntPtr.Zero);
                if (result != 1)
                {
                    throw new Exception("SetWindowSubClass Failed");
                }
                Debug.WriteLine("SubclassingWindow.AssignHandle called for hWnd " + Hwnd);
                //DO NOT REMOVE THIS CALL. Dockable windows are instantiated by the VBE, not directly by RD.  On top of that,
                //since we have to inherit from UserControl we don't have to keep handling window messages until the VBE gets
                //around to destroying the control's host or it results in an access violation when the base class is disposed.
                //We need to manually call base.Dispose() ONLY in response to a WM_DESTROY message.
                _thisHandle = GCHandle.Alloc(this, GCHandleType.Normal);
                _listening = true;
            }
        }

        private void ReleaseHandle()
        {
            lock (_subclassLock)
            {
                if (!_listening)
                {
                    return;
                }
                Debug.WriteLine("SubclassingWindow.ReleaseHandle called for hWnd " + Hwnd);
                var result = RemoveWindowSubclass(Hwnd, _wndProc, _subclassId);
                if (result != 1)
                {
                    throw new Exception("RemoveWindowSubclass Failed");
                }
                _listening = false;
            }
        }

        protected virtual void HandleResized(int width, int height)
        {
            // no-op default
        }

        protected virtual int SubClassProc(IntPtr hWnd, IntPtr msg, IntPtr wParam, IntPtr lParam, IntPtr uIdSubclass, IntPtr dwRefData)
        {
            if (!_listening)
            {
                Debug.WriteLine("State corrupted. Received window message while not listening.");
                return DefSubclassProc(hWnd, msg, wParam, lParam);
            }

            //if ((uint)msg == (uint)WM.SIZE)
            //{
            //    HandleResized(lParam.LoWord(), lParam.HiWord());
            //}

            if ((uint) msg == (uint) WM.WINDOWPOSCHANGED)
            {
                var pos = (WindowPos)Marshal.PtrToStructure(lParam, typeof(WindowPos));
                HandleResized(pos.cx, pos.cy);
            }

            if ((uint)msg == (uint)WM.RUBBERDUCK_SINKING || (uint)msg == (uint)WM.DESTROY)
            {
                ReleaseHandle();                
            }
            return DefSubclassProc(hWnd, msg, wParam, lParam);
        }
    }
}