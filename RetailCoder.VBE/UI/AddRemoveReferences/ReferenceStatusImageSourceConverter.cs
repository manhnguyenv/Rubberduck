using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Media;
using Rubberduck.AddRemoveReferences;
using ImageSourceConverter = Rubberduck.UI.Converters.ImageSourceConverter;

namespace Rubberduck.UI.AddRemoveReferences
{
    public class ReferenceStatusImageSourceConverter : ImageSourceConverter
    {
        private readonly IDictionary<ReferenceStatus, ImageSource> _icons =
            new Dictionary<ReferenceStatus, ImageSource>
            {
                { ReferenceStatus.None, null },
                { ReferenceStatus.BuiltIn, ToImageSource(Properties.Resources.padlock) },
                { ReferenceStatus.Broken, ToImageSource(Properties.Resources.exclamation_diamond) },
                { ReferenceStatus.Loaded, ToImageSource(Properties.Resources.tick) },
            };

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || value.GetType() != typeof(ReferenceStatus))
            {
                return null;
            }

            var status = (ReferenceStatus)value;
            return _icons[status];
        }
    }
}