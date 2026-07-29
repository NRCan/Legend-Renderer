using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Media;

namespace GSCLegendRendererPro.Converters
{
    public class ColorFromBoolConverter: IValueConverter
    {
        public ColorFromBoolConverter()
        {
        }

        object IValueConverter.Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool && (bool)value)
            {
                //Set as black
                SolidColorBrush outColor = new SolidColorBrush();
                outColor.Color = Color.FromRgb(0, 0, 0);

                return outColor;
            }
            else
            {
                //Set as red
                SolidColorBrush outColor = new SolidColorBrush();
                outColor.Color = Color.FromRgb(255, 0, 0);

                return outColor;
            }

        }

        object IValueConverter.ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Color && (Color)value == Color.FromRgb(0, 0, 0))
            {
                return true;
            }
            return false;
        }
    }
}
