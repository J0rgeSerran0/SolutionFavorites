using SolutionFavorites.MEF;

namespace SolutionFavorites.Commands
{
    /// <summary>
    /// Command to remove a folder from favorites.
    /// </summary>
    [Command(PackageIds.RemoveFolder)]
    internal sealed class RemoveFolderCommand : BaseCommand<RemoveFolderCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var currentItem = FavoritesContextMenuController.CurrentItem;

            if (currentItem is FavoriteFolderNode folderNode)
            {
                FavoritesManager.Instance.Remove(folderNode.Item);
            }
        }
    }
}
