using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Data;

namespace TournamentFormGenerator.Utils
{
    public class BoolToOppositeBoolConverter : IValueConverter
    {
        #region IValueConverter Members

        public object Convert(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            if (targetType != typeof(bool) && targetType!=typeof(bool?))
                throw new InvalidOperationException("The target must be a boolean");

            if (value == null)
                return null;
            else
                return !(bool)value;
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            return Convert(value, targetType, parameter, culture);
        }

        #endregion
    }
}
