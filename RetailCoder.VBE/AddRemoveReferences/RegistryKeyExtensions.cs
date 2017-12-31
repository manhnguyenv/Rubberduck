using System;
using Microsoft.Win32;

namespace Rubberduck.AddRemoveReferences
{
    public static class RegistryKeyExtensions
    {
        /// <summary>
        /// Gets the name of a key, without its parent/path.
        /// </summary>
        public static string GetKeyName(this RegistryKey key)
        {
            var name = key?.Name;
            return name?.Substring(name.LastIndexOf(@"\", StringComparison.InvariantCultureIgnoreCase) + 1) ?? string.Empty;
        }
    }
}