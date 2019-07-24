using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Threading;
using TemplatesTool;
using TemplatesTool.Models;
using WPF.Themes;
using Application = System.Windows.Application;
using Img = System.Windows.Controls.Image;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.Forms.MessageBox;
using TextBox = System.Windows.Controls.TextBox;
using TreeView = System.Windows.Controls.TreeView;

namespace templatesUI
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.ApplyTheme("ExpressionDark");
        }

        private void Window_ContentRendered(object sender, EventArgs e)
        {
            Finder.FindFirstResult<Img>(this, "RefreshImage").Source =
                ItemCreator.Bitmap2BitmapSource(Properties.Resources.refresh);
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            var textbox = Finder.FindFirstVisibleResult<TextBox>(this, "BrowseTextBox");
            var path = BrowseDirectory();
            if (path == null) return;
            textbox.Text = path;
            UpdateTreeView();
        }

        private string BrowseDirectory()
        {
            var fbd = new FolderBrowserDialog {SelectedPath = @"\\"};
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK) return fbd.SelectedPath;
            return null;
        }

        private void RefreshBrowseButton_Click(object sender, RoutedEventArgs e)
        {
            UpdateTreeView();
        }

        private void BrowseTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter) return;
            UpdateTreeView();
        }

        private void UpdateTreeView()
        {
            var treeView = BrowseTreeView;
            var textBox = Finder.FindFirstResult<TextBox>(this, "BrowseTextBox");

            treeView.Items.Clear();
            try
            {
                var directoryInfo = new DirectoryInfo(textBox.Text);
                if (directoryInfo.Exists)
                {
                    var dirItem = Tree.CreateDirectoryItem(treeView.Items, directoryInfo);
                    Tree.RecursivelyCreateTreeView(dirItem, directoryInfo);
                    dirItem.ExpandSubtree();
                }
            }
            catch
            {
            }

            Application.Current.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, new Action(() => { }))
                .Wait();

            foreach (var tmpObj in Finder.Find<ScrollViewer>(treeView))
                tmpObj.ScrollToHome();
        }

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var treeView = sender as TreeView;
            var selectedItem = treeView.SelectedItem as TreeViewItem;
            Tree.SelectionChanged(selectedItem);
        }

        private async void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            DisableUI();
            var fileItems = GetSelectedItems(BrowseTreeView.Items);
            foreach (var file in fileItems)
            {
                SetItemToPending(file);
                file.ToolTip = null;
            }

            foreach (var file in fileItems)
            {
                if (IsFileDirectory(file))
                {
                    if (file.Items.Count == 0) SetItemToSuccess(file);
                    continue;
                }

                if (file.Items.Count == 0) SetItemToInProgress(file);

                var fileInfo = ((ItemTag) file.Tag).FileInfo;

                var s = ((ItemTag) file.Tag).MicrosoftItem;
                var result = await Task.Run(() => ResolveNode(s, fileInfo));
                if (result == "true")
                    SetItemToSuccess(file);
                else
                    SetItemToError(file, result);
            }

            //try this
            //string rootDir = BrowseTextBox.Text;

            //Write an complete log file under root folder
            var rootDir = Finder.FindFirstResult<TextBox>(this, "BrowseTextBox").Text;
            Common.WriteAllText(rootDir); // Not working currently

            MessageBox.Show(@"Finished!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

            EnableUI();
        }

        private void EnableUI()
        {
            BrowseTextBoxParent.IsEnabled = true;
            ApplyButton.IsEnabled = true;
            BrowseButton.IsEnabled = true;
            foreach (TreeViewItem item in BrowseTreeView.Items)
                item.IsEnabled = true;
        }

        private void DisableUI()
        {
            BrowseTextBoxParent.IsEnabled = false;
            ApplyButton.IsEnabled = false;
            BrowseButton.IsEnabled = false;
            foreach (TreeViewItem item in BrowseTreeView.Items)
                item.IsEnabled = false;
        }


        private List<TreeViewItem> GetSelectedItems(ItemCollection Items)
        {
            var selectedItems = new List<TreeViewItem>();

            foreach (TreeViewItem childItem in Items)
            {
                var tag = (ItemTag) childItem.Tag;

                if (tag.Selected)
                {
                    selectedItems.Add(childItem);
                    selectedItems.AddRange(GetSelectedItems(childItem.Items));
                }
            }

            return selectedItems;
        }

        private bool IsFileDirectory(TreeViewItem file)
        {
            var tag = (ItemTag) file.Tag;
            return tag.MicrosoftItem == MicrosoftItem.directory;
        }

        private void SetItemToPending(TreeViewItem item)
        {
            var itemTag = item.Tag as ItemTag;
            var img = Finder.FindFirstResult<Img>(item, "image");

            itemTag.Color = "orange";
            img.Source =
                ItemCreator.Bitmap2BitmapSource(
                    (Bitmap) Properties.Resources.ResourceManager.GetObject(itemTag.MicrosoftItem + "_orange"));
        }

        private void SetItemToInProgress(TreeViewItem item)
        {
            var itemTag = item.Tag as ItemTag;
            var img = Finder.FindFirstResult<Img>(item, "image");

            itemTag.Color = "yellow";
            img.Source =
                ItemCreator.Bitmap2BitmapSource(
                    (Bitmap) Properties.Resources.ResourceManager.GetObject(itemTag.MicrosoftItem + "_yellow"));
        }

        private void SetItemToError(TreeViewItem item, string message)
        {
            if (!IsFileDirectory(item))
                item.ToolTip = $"Error: {message}";

            var itemTag = item.Tag as ItemTag;
            var img = Finder.FindFirstResult<Img>(item, "image");

            itemTag.Color = "red";
            img.Source =
                ItemCreator.Bitmap2BitmapSource(
                    (Bitmap) Properties.Resources.ResourceManager.GetObject(itemTag.MicrosoftItem + "_red"));

            if (!(item.Parent is TreeView))
                SetParentItemState((TreeViewItem) item.Parent);
        }

        private void SetItemToSuccess(TreeViewItem item)
        {
            var itemTag = item.Tag as ItemTag;
            var img = Finder.FindFirstResult<Img>(item, "image");

            itemTag.Color = "green";
            img.Source =
                ItemCreator.Bitmap2BitmapSource(
                    (Bitmap) Properties.Resources.ResourceManager.GetObject(itemTag.MicrosoftItem + "_green"));

            if (!(item.Parent is TreeView))
                SetParentItemState((TreeViewItem) item.Parent);
        }

        private void SetParentItemState(TreeViewItem item)
        {
            var containsError = false;
            foreach (TreeViewItem subitem in item.Items)
            {
                if (((ItemTag) subitem.Tag).Color == "orange") return;
                if (((ItemTag) subitem.Tag).Color == "red") containsError = true;
            }

            if (containsError)
                SetItemToError(item, null);
            else
                SetItemToSuccess(item);
        }

        public string ResolveNode(MicrosoftItem msItem, FileSystemInfo file)
        {
            switch (msItem)
            {
                case MicrosoftItem.excel:
                {
                    return ResolveExcel(file);
                }

                case MicrosoftItem.word:
                {
                    return ResolveWord(file);
                }

                case MicrosoftItem.powerpoint:
                {
                    return ResolvePowerpoint(file);
                }
            }

            return string.Empty;
        }

        private string ResolvePowerpoint(FileSystemInfo file)
        {
            //"if success -> return "true", else return error message string

            var filePath = file.FullName;
            var objPpt = new PowerPointHandler(filePath);

            try
            {
                objPpt.PptMain(filePath);
            }
            catch (Exception e)
            {
                return e.Message;
            }
            finally
            {
                Common.WriteSglText(filePath);
            }

            return "true";
            throw new NotImplementedException();
        }

        private string ResolveWord(FileSystemInfo file)
        {
            //"if success -> return "true", else return error message string

            var filePath = file.FullName;
            var objWordHandler = new WordHandler(filePath);

            try
            {
                objWordHandler.WordMain(filePath);
            }
            catch (Exception e)
            {
                return e.Message;
            }
            finally
            {
                Common.WriteSglText(filePath);
            }

            return "true";
            throw new NotImplementedException();
        }


        private string ResolveExcel(FileSystemInfo file)
        {
            //"if success -> return "true", else return error message string

            var filePath = file.FullName;
            var objExcelHandler = new ExcelHandler(filePath);

            try
            {
                objExcelHandler.ExcelMain(filePath);
            }
            catch (Exception e)
            {
                return e.Message;
            }
            finally
            {
                Common.WriteSglText(filePath);
            }

            return "true";
            throw new NotImplementedException();
        }
    }
}