using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using Microsoft.Internal.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Imaging.Interop;
using SolutionFavorites.Models;

namespace SolutionFavorites.MEF
{
    /// <summary>
    /// Represents a virtual folder node in the Favorites tree.
    /// </summary>
    internal sealed class FavoriteFolderNode :
        IAttachedCollectionSource,
        ITreeDisplayItem,
        ITreeDisplayItemWithImages,
        IPrioritizedComparable,
        IBrowsablePattern,
        IInteractionPatternProvider,
        IContextMenuPattern,
        IInvocationPattern,
        IDragDropSourcePattern,
        IDragDropTargetPattern,
        ISupportDisposalNotification,
        INotifyPropertyChanged,
        IDisposable
    {
        private readonly ObservableCollection<object> _children;
        private bool _disposed;
        private bool _isExpanded;

        private static readonly HashSet<Type> _supportedPatterns = new HashSet<Type>
        {
            typeof(ITreeDisplayItem),
            typeof(IBrowsablePattern),
            typeof(IContextMenuPattern),
            typeof(IInvocationPattern),
            typeof(IDragDropSourcePattern),
            typeof(IDragDropTargetPattern),
            typeof(ISupportDisposalNotification),
        };

        public FavoriteFolderNode(FavoriteItem item, object parent)
        {
            Item = item;
            SourceItem = parent;
            _children = new ObservableCollection<object>();
            FavoritesManager.Instance.FavoritesChanged += OnFavoritesChanged;
            RefreshChildren();
        }

        /// <summary>
        /// The underlying favorite folder item.
        /// </summary>
        public FavoriteItem Item { get; }

        /// <summary>
        /// Parent node.
        /// </summary>
        public object SourceItem { get; }

        private void OnFavoritesChanged(object sender, EventArgs e)
        {
#pragma warning disable VSTHRD110 // Observe result of async calls
            _ = ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                RefreshChildren();
            });
#pragma warning restore VSTHRD110
        }

        /// <summary>
        /// Refreshes the children of this folder.
        /// </summary>
        public void RefreshChildren()
        {
            // Dispose existing children
            foreach (var child in _children)
            {
                (child as IDisposable)?.Dispose();
            }
            _children.Clear();

            var folderItems = FavoritesManager.Instance.GetFolderItems(Item);
            foreach (var item in folderItems)
            {
                if (item.IsFolder)
                {
                    _children.Add(new FavoriteFolderNode(item, this));
                }
                else
                {
                    _children.Add(new FavoriteFileNode(item, this));
                }
            }

            RaisePropertyChanged(nameof(HasItems));
            RaisePropertyChanged(nameof(Items));
        }

        // IAttachedCollectionSource
        public bool HasItems => Item.Children != null && Item.Children.Count > 0;
        public IEnumerable Items => _children;

        // ITreeDisplayItem
        public string Text => Item.Name;
        public string ToolTipText => Item.Name;
        public object ToolTipContent => ToolTipText;
        public string StateToolTipText => string.Empty;
        System.Windows.FontWeight ITreeDisplayItem.FontWeight => System.Windows.FontWeights.Normal;
        System.Windows.FontStyle ITreeDisplayItem.FontStyle => System.Windows.FontStyles.Normal;
        public bool IsCut => false;

        // ITreeDisplayItemWithImages
        public ImageMoniker IconMoniker => _isExpanded ? KnownMonikers.FolderOpened : KnownMonikers.FolderClosed;
        public ImageMoniker ExpandedIconMoniker => KnownMonikers.FolderOpened;
        public ImageMoniker OverlayIconMoniker => default;
        public ImageMoniker StateIconMoniker => default;

        // IPrioritizedComparable - Folders appear before files
        public int Priority => 0;

        public int CompareTo(object obj)
        {
            if (obj is FavoriteFolderNode otherFolder)
            {
                return StringComparer.OrdinalIgnoreCase.Compare(Text, otherFolder.Text);
            }
            if (obj is FavoriteFileNode)
            {
                return -1; // Folders before files
            }
            return 0;
        }

        // IBrowsablePattern
        public object GetBrowseObject() => this;

        // IInteractionPatternProvider
        public TPattern GetPattern<TPattern>() where TPattern : class
        {
            if (!_disposed && _supportedPatterns.Contains(typeof(TPattern)))
            {
                return this as TPattern;
            }

            if (typeof(TPattern) == typeof(ISupportDisposalNotification))
            {
                return this as TPattern;
            }


            return null;
        }

        // IContextMenuPattern
        public IContextMenuController ContextMenuController => FavoritesContextMenuController.Instance;

        // IInvocationPattern - double-click expands/collapses folder
        public IInvocationController InvocationController => null;
        public bool CanPreview => false;

        // IDragDropSourcePattern
        public IDragDropSourceController DragDropSourceController => FavoritesDragDropController.Instance;

        // IDragDropTargetPattern
        public DirectionalDropArea SupportedAreas => DirectionalDropArea.On;

        public void OnDragEnter(DirectionalDropArea dropArea, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(FavoritesDragDropTargetController.FavoritesDataFormat))
            {
                e.Effects = DragDropEffects.Move;
            }
        }

        public void OnDragOver(DirectionalDropArea dropArea, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(FavoritesDragDropTargetController.FavoritesDataFormat))
            {
                e.Effects = DragDropEffects.Move;
            }
        }

        public void OnDragLeave(DirectionalDropArea dropArea, DragEventArgs e)
        {
            e.Effects = DragDropEffects.None;
        }

        public void OnDrop(DirectionalDropArea dropArea, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(FavoritesDragDropTargetController.FavoritesDataFormat))
            {
                var nodes = e.Data.GetData(FavoritesDragDropTargetController.FavoritesDataFormat) as object[];
                if (nodes != null)
                {
                    foreach (var node in nodes)
                    {
                        Models.FavoriteItem itemToMove = null;
                        
                        if (node is FavoriteFileNode fileNode)
                            itemToMove = fileNode.Item;
                        else if (node is FavoriteFolderNode folderNode && folderNode.Item != Item)
                            itemToMove = folderNode.Item;

                        if (itemToMove != null)
                        {
                            FavoritesManager.Instance.MoveItem(itemToMove, Item);
                        }
                    }
                    e.Handled = true;
                }
            }
        }

        // ISupportDisposalNotification
        public bool IsDisposed => _disposed;

        // INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Updates the expanded state for icon changes.
        /// </summary>
        public void SetExpanded(bool expanded)
        {
            if (_isExpanded != expanded)
            {
                _isExpanded = expanded;
                RaisePropertyChanged(nameof(IconMoniker));
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                FavoritesManager.Instance.FavoritesChanged -= OnFavoritesChanged;
                
                foreach (var child in _children)
                {
                    (child as IDisposable)?.Dispose();
                }
                _children.Clear();
                
                RaisePropertyChanged(nameof(IsDisposed));
            }
        }
    }
}
