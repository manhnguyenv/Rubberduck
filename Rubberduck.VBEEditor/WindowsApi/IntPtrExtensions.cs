using System;

namespace Rubberduck.VBEditor.WindowsApi
{
    public static class IntPtrExtensions
    {
        public static int LoWord(this IntPtr value)
        {
            return unchecked((short) (long) value);
        }

        public static int HiWord(this IntPtr value)
        {
            return unchecked((short) ((long) value >> 16));
        }
    }
}