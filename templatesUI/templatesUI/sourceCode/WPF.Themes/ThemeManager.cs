﻿using System;
using System.Windows;
using System.Windows.Controls;

namespace WPF.Themes
{
    public static class ThemeManager
    {
        public static ResourceDictionary GetThemeResourceDictionary(string theme)
        {
            if (theme != null)
            {
                //Assembly assembly = Assembly.LoadFrom("WPF.Themes.dll");
                var packUri = string.Format(@"/WPF.Themes;component/{0}/Theme.xaml", theme);
                return Application.LoadComponent(new Uri(packUri, UriKind.Relative)) as ResourceDictionary;
            }

            return null;
        }

        public static string[] GetThemes()
        {
            return new[]
            {
                "ExpressionDark", "ExpressionLight",
                //"RainierOrange", "RainierPurple", "RainierRadialBlue", 
                "ShinyBlue", "ShinyRed",
                //"ShinyDarkTeal", "ShinyDarkGreen", "ShinyDarkPurple",
                "DavesGlossyControls",
                "WhistlerBlue",
                "BureauBlack", "BureauBlue",
                "BubbleCreme",
                "TwilightBlue",
                "UXMusingsRed", "UXMusingsGreen",
                //"UXMusingsRoughRed", "UXMusingsRoughGreen", 
                "UXMusingsBubblyBlue"
            };
        }

        public static void ApplyTheme(this Application app, string theme)
        {
            var dictionary = GetThemeResourceDictionary(theme);

            if (dictionary != null)
            {
                app.Resources.MergedDictionaries.Clear();
                app.Resources.MergedDictionaries.Add(dictionary);
            }
        }

        public static void ApplyTheme(this ContentControl control, string theme)
        {
            var dictionary = GetThemeResourceDictionary(theme);

            if (dictionary != null)
            {
                control.Resources.MergedDictionaries.Clear();
                control.Resources.MergedDictionaries.Add(dictionary);
            }
        }

        #region Theme

        /// <summary>
        ///     Theme Attached Dependency Property
        /// </summary>
        public static readonly DependencyProperty ThemeProperty =
            DependencyProperty.RegisterAttached("Theme", typeof(string), typeof(ThemeManager),
                new FrameworkPropertyMetadata(string.Empty,
                    OnThemeChanged));

        /// <summary>
        ///     Gets the Theme property.  This dependency property
        ///     indicates ....
        /// </summary>
        public static string GetTheme(DependencyObject d)
        {
            return (string) d.GetValue(ThemeProperty);
        }

        /// <summary>
        ///     Sets the Theme property.  This dependency property
        ///     indicates ....
        /// </summary>
        public static void SetTheme(DependencyObject d, string value)
        {
            d.SetValue(ThemeProperty, value);
        }

        /// <summary>
        ///     Handles changes to the Theme property.
        /// </summary>
        private static void OnThemeChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var theme = e.NewValue as string;
            if (theme == string.Empty)
                return;

            var control = d as ContentControl;
            if (control != null) control.ApplyTheme(theme);
        }

        #endregion
    }
}