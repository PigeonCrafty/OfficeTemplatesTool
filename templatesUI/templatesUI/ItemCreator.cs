using System;
using System.Drawing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using templatesUI.Properties;
using Brushes = System.Windows.Media.Brushes;
using Image = System.Windows.Controls.Image;

namespace templatesUI
{
    internal class ItemCreator
    {
        public static TextBlock Create(string name, MicrosoftItem msItem)
        {
            var tempTextBlock = new TextBlock();
            tempTextBlock.Height = 16;
            var tempImage = new Image();

            tempImage.Margin = new Thickness(1, 1, 5, 1);
            tempImage.Name = "image";

            tempImage.Source = Bitmap2BitmapSource((Bitmap) Resources.ResourceManager.GetObject(msItem + "_white"));

            var lab = new Label
            {
                Content = name,
                Margin = new Thickness(0, 0, 0, 0),
                Foreground = Brushes.White,
                Name = "label"
            };

            tempTextBlock.Inlines.Add(tempImage);
            tempTextBlock.Inlines.Add(lab);

            return tempTextBlock;
        }

        public static BitmapSource Bitmap2BitmapSource(Bitmap bitmap)
        {
            var i = Imaging.CreateBitmapSourceFromHBitmap(
                bitmap.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
            return i;
        }
    }
}