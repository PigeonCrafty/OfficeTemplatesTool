using System;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using templatesUI.Properties;
using Brushes = System.Windows.Media.Brushes;
using Image = System.Windows.Controls.Image;

namespace templatesUI
{
    internal class Tree
    {
        public static bool RecursivelyCreateTreeView(TreeViewItem rootItem, DirectoryInfo directory)
        {
            foreach (var subDirectory in directory.GetDirectories())
            {
                var dirItem = CreateDirectoryItem(rootItem.Items, subDirectory);
                if (!RecursivelyCreateTreeView(dirItem, subDirectory))
                    rootItem.Items.Remove(dirItem);
            }

            foreach (var file in directory.GetFiles())
                CreateFileNode(file, rootItem);

            return rootItem.Items.Count != 0;
        }

        public static TreeViewItem CreateDirectoryItem(ItemCollection treeItemCollection, DirectoryInfo directoryInfo)
        {
            var msItem = MicrosoftItem.directory;
            var item = new TreeViewItem
            {
                Header = ItemCreator.Create(directoryInfo.Name, msItem),
                Tag = new ItemTag
                {
                    FileInfo = directoryInfo,
                    MicrosoftItem = msItem,
                    Color = "white"
                }
            };
            treeItemCollection.Add(item);
            return item;
        }

        private static void CreateFileNode(FileInfo file, TreeViewItem rootItem)
        {
            if (file.Name.StartsWith("~$")) return;
            switch (ChooseApp(file.Extension))
            {
                case MicrosoftItem.word:
                {
                    CreateMicrosoftItem(file, rootItem, MicrosoftItem.word);
                    break;
                }
                case MicrosoftItem.excel:
                {
                    CreateMicrosoftItem(file, rootItem, MicrosoftItem.excel);
                    break;
                }
                case MicrosoftItem.powerpoint:
                {
                    CreateMicrosoftItem(file, rootItem, MicrosoftItem.powerpoint);
                    break;
                }
            }
        }

        public static MicrosoftItem ChooseApp(string extension)
        {
            switch (extension)
            {
                case ".potx":
                case ".pptx":
                    return MicrosoftItem.powerpoint;
                case ".dotx":
                case ".docx":
                    return MicrosoftItem.word;
                case ".xltx":
                case ".xlsx":
                case ".xltm":
                    return MicrosoftItem.excel;
                default:
                    return MicrosoftItem.directory;
            }
        }

        private static void CreateMicrosoftItem(FileInfo file, TreeViewItem rootItem, MicrosoftItem microsoftItem)
        {
            var item = new TreeViewItem
            {
                Header = ItemCreator.Create(file.Name, microsoftItem),
                Tag = new ItemTag
                {
                    FileInfo = file,
                    MicrosoftItem = microsoftItem,
                    Color = "white"
                }
            };
            rootItem.Items.Add(item);
        }

        public static void SelectionChanged(TreeViewItem selectedItem)
        {
            if (selectedItem == null)
                return;
            selectedItem.IsSelected = false;

            var itemTag = (ItemTag) selectedItem.Tag;
            itemTag.Selected = !itemTag.Selected;
            selectedItem.Tag = itemTag;
            ChangeGraphicAfterSelection(selectedItem);
            CheckChildNodes(selectedItem);
            CheckParentNode(selectedItem);
        }

        private static void CheckChildNodes(TreeViewItem item)
        {
            var itemTag = (ItemTag) item.Tag;
            foreach (TreeViewItem childItem in item.Items)
            {
                var childItemTag = (ItemTag) childItem.Tag;

                var change = childItemTag.Selected != itemTag.Selected;
                childItemTag.Selected = itemTag.Selected;

                childItem.Tag = childItemTag;

                if (change)
                    ChangeGraphicAfterSelection(childItem);

                CheckChildNodes(childItem);
            }
        }

        private static void CheckParentNode(TreeViewItem item)
        {
            if (item.Parent is TreeView) return;

            var select = false;
            foreach (TreeViewItem childItem in ((TreeViewItem) item.Parent).Items)
            {
                var childItemTag = (ItemTag) childItem.Tag;
                if (childItemTag.Selected)
                {
                    select = true;
                    break;
                }
            }

            var parentItemTag = (ItemTag) ((TreeViewItem) item.Parent).Tag;

            var change = parentItemTag.Selected != select;
            parentItemTag.Selected = select;
            ((TreeViewItem) item.Parent).Tag = parentItemTag;

            if (change)
                ChangeGraphicAfterSelection((TreeViewItem) item.Parent);

            CheckParentNode((TreeViewItem) item.Parent);
        }

        private static void ChangeGraphicAfterSelection(TreeViewItem selectedItem)
        {
            var itemTag = (ItemTag) selectedItem.Tag;

            if (itemTag.Selected)
            {
                var border = Finder.FindFirstResult<Border>(selectedItem, "SelectionBorder");

                var fadeIn = new DoubleAnimationUsingKeyFrames();
                fadeIn.BeginTime = TimeSpan.FromSeconds(0);
                fadeIn.KeyFrames.Add(new SplineDoubleKeyFrame(1,
                    new TimeSpan(0, 0, 0, 0, 100)
                ));

                var sb = new Storyboard();
                Storyboard.SetTarget(fadeIn, border);
                Storyboard.SetTargetProperty(fadeIn, new PropertyPath("(Opacity)"));
                sb.Children.Add(fadeIn);
                selectedItem.Resources.Clear();
                selectedItem.Resources.Add("MyEffect", sb);

                sb.Begin();

                var lab = Finder.FindFirstResult<Label>(selectedItem, "label");
                lab.Foreground = Brushes.Black;

                if (itemTag.Color == "white")
                {
                    itemTag.Color = "black";
                    var img = Finder.FindFirstResult<Image>(selectedItem, "image");
                    img.Source =
                        ItemCreator.Bitmap2BitmapSource(
                            (Bitmap) Resources.ResourceManager.GetObject(itemTag.MicrosoftItem + "_black"));
                }
            }
            else
            {
                var border = Finder.FindFirstResult<Border>(selectedItem, "SelectionBorder");

                var fadeIn = new DoubleAnimationUsingKeyFrames();
                fadeIn.BeginTime = TimeSpan.FromSeconds(0);
                fadeIn.KeyFrames.Add(new SplineDoubleKeyFrame(0,
                    new TimeSpan(0, 0, 0, 0, 200)
                ));

                var sb = new Storyboard();
                Storyboard.SetTarget(fadeIn, border);
                Storyboard.SetTargetProperty(fadeIn, new PropertyPath("(Opacity)"));
                sb.Children.Add(fadeIn);
                selectedItem.Resources.Clear();
                selectedItem.Resources.Add("MyEffect", sb);

                sb.Begin();

                var lab = Finder.FindFirstResult<Label>(selectedItem, "label");
                lab.Foreground = Brushes.White;

                if (itemTag.Color == "black")
                {
                    itemTag.Color = "white";
                    var img = Finder.FindFirstResult<Image>(selectedItem, "image");
                    img.Source =
                        ItemCreator.Bitmap2BitmapSource(
                            (Bitmap) Resources.ResourceManager.GetObject(itemTag.MicrosoftItem + "_white"));
                }
            }
        }
    }
}