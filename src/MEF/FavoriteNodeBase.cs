using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using Microsoft.Internal.VisualStudio.PlatformUI;
using SolutionFavorites.Models;

namespace SolutionFavorites.MEF
{
    /// <summary>
    /// Base class for all favorite tree nodes, providing shared functionality for
    /// disposal notification, property change notification, and pattern support.
    /// </summary>
    internal abstract class FavoriteNodeBase :
        ITreeDisplayItem,
        IBrowsablePattern,
        IInteractionPatternProvider,
        IContextMenuPattern,
        ISupportDisposalNotification,
        INotifyPropertyChanged,
        IDisposable
    {
        private bool _disposed;

        /// <summary>
        /// The set of pattern types this node supports.
        /// Derived classes should override to add additional patterns.
        /// </summary>
        protected virtual HashSet<Type> SupportedPatterns { get; } = new HashSet<Type>
        {
            typeof(ITreeDisplayItem),
            typeof(IBrowsablePattern),
            typeof(IContextMenuPattern),
            typeof(ISupportDisposalNotification),
        };

        /// <summary>
        /// Parent node in the tree.
        /// </summary>
        public object SourceItem { get; }

        protected FavoriteNodeBase(object parent)
        {
            SourceItem = parent;
        }

        // ITreeDisplayItem - abstract properties for derived classes to implement
        public abstract string Text { get; }
        public virtual string ToolTipText => Text;
        public virtual object ToolTipContent => ToolTipText;
        public virtual string StateToolTipText => string.Empty;
        public virtual FontWeight FontWeight => FontWeights.Normal;
        System.Windows.FontStyle ITreeDisplayItem.FontStyle => System.Windows.FontStyles.Normal;
        public virtual bool IsCut => false;

        // IBrowsablePattern
        public object GetBrowseObject() => this;

        // IContextMenuPattern
        public IContextMenuController ContextMenuController => FavoritesContextMenuController.Instance;

        // IInteractionPatternProvider
        public virtual TPattern GetPattern<TPattern>() where TPattern : class
        {
            if (!_disposed && SupportedPatterns.Contains(typeof(TPattern)))
            {
                return this as TPattern;
            }

            if (typeof(TPattern) == typeof(ISupportDisposalNotification))
            {
                return this as TPattern;
            }

            return null;
        }

        // ISupportDisposalNotification
        public bool IsDisposed => _disposed;

        // INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        protected void RaisePropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Called when the node is being disposed. Override to add cleanup logic.
        /// </summary>
        protected virtual void OnDisposing() { }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                OnDisposing();
                RaisePropertyChanged(nameof(IsDisposed));
            }
        }

        #region Drag-Drop Target Helpers

        /// <summary>
        /// Checks if the drag data contains items that can be added to favorites.
        /// </summary>
        private static bool CanAcceptDrop(DragEventArgs e)
        {
            // Accept our internal favorites format (for reordering)
            if (e.Data.GetDataPresent(FavoritesDragDropConstants.FavoritesDataFormat))
            {
                return true;
            }

            // Accept file drops from Solution Explorer or Windows Explorer
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Handles drag enter for favorites items.
        /// </summary>
        protected static void HandleDragEnter(DragEventArgs e)
        {
            if (CanAcceptDrop(e))
            {
                e.Effects = e.Data.GetDataPresent(FavoritesDragDropConstants.FavoritesDataFormat)
                    ? DragDropEffects.Move
                    : DragDropEffects.Copy;
            }
        }

        /// <summary>
        /// Handles drag over for favorites items.
        /// </summary>
        protected static void HandleDragOver(DragEventArgs e)
        {
            if (CanAcceptDrop(e))
            {
                e.Effects = e.Data.GetDataPresent(FavoritesDragDropConstants.FavoritesDataFormat)
                    ? DragDropEffects.Move
                    : DragDropEffects.Copy;
            }
        }

        /// <summary>
        /// Handles drag leave for favorites items.
        /// </summary>
        protected static void HandleDragLeave(DragEventArgs e)
        {
            e.Effects = DragDropEffects.None;
        }

        /// <summary>
        /// Handles drop for favorites items, moving them to the target folder.
        /// </summary>
        /// <param name="targetFolder">The target folder, or null for root level.</param>
        /// <param name="e">The drag event args.</param>
        protected static void HandleDrop(FavoriteItem targetFolder, DragEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Handle internal favorites reordering
            if (e.Data.GetDataPresent(FavoritesDragDropConstants.FavoritesDataFormat))
            {
                HandleFavoritesDrop(targetFolder, e);
                return;
            }

            // Handle file drops from Solution Explorer or Windows Explorer
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                HandleFileDrop(targetFolder, e);
                return;
            }
        }

        /// <summary>
        /// Handles dropping favorites items for reordering.
        /// </summary>
        private static void HandleFavoritesDrop(FavoriteItem targetFolder, DragEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var nodes = e.Data.GetData(FavoritesDragDropConstants.FavoritesDataFormat) as object[];
            if (nodes == null)
                return;

            foreach (var node in nodes)
            {
                FavoriteItem itemToMove = null;

                if (node is FavoriteFileNode fileNode)
                {
                    itemToMove = fileNode.Item;
                }
                else if (node is FavoriteFolderNode folderNode)
                {
                    // Don't allow dropping a folder onto itself
                    if (targetFolder != null && folderNode.Item == targetFolder)
                        continue;
                    itemToMove = folderNode.Item;
                }

                if (itemToMove != null)
                {
                    FavoritesManager.Instance.MoveItem(itemToMove, targetFolder);
                }
            }

            e.Handled = true;
        }

        /// <summary>
        /// Handles dropping files from Solution Explorer or Windows Explorer.
        /// </summary>
        private static void HandleFileDrop(FavoriteItem targetFolder, DragEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0)
                return;

            foreach (var filePath in files)
            {
                // Only add files, not directories
                if (System.IO.File.Exists(filePath))
                {
                    if (targetFolder != null)
                    {
                        FavoritesManager.Instance.AddFileToFolder(filePath, targetFolder);
                    }
                    else
                    {
                        FavoritesManager.Instance.AddFile(filePath);
                    }
                }
            }

            e.Handled = true;
        }

        #endregion

        #region Child Collection Helpers

        /// <summary>
        /// Disposes all children in a collection and clears it.
        /// </summary>
        protected static void DisposeChildren(ObservableCollection<object> children)
        {
            foreach (var child in children)
            {
                (child as IDisposable)?.Dispose();
            }
            children.Clear();
        }

        /// <summary>
        /// Creates the appropriate node type for a favorite item.
        /// </summary>
        protected static object CreateNodeForItem(FavoriteItem item, object parent)
        {
            return item.IsFolder
                ? (object)new FavoriteFolderNode(item, parent)
                : new FavoriteFileNode(item, parent);
        }

        #endregion
    }
}
