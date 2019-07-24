using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;

namespace templatesUI
{
    public static class Finder
    {
        public static T FindFirstResult<T>(FrameworkElement depObj, string name) where T : FrameworkElement
        {
            foreach (var tmpObj in Find<T>(depObj))
                if (tmpObj.Name == name)
                    return tmpObj;
            return null;
        }

        public static T FindFirstVisibleResult<T>(FrameworkElement depObj, string name) where T : FrameworkElement
        {
            foreach (var tmpObj in Find<T>(depObj))
                if (tmpObj.Name == name)
                    if (tmpObj.IsVisible)
                        return tmpObj;
            return null;
        }

        public static IEnumerable<T> Find<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
                for (var i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    var child = VisualTreeHelper.GetChild(depObj, i);

                    if (child != null && child is T)
                        yield return (T) child;

                    foreach (var childOfChild in Find<T>(child))
                        yield return childOfChild;
                }
        }
    }
}