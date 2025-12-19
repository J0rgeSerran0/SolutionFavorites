using SolutionFavorites.MEF;

namespace SolutionFavorites.Commands
{
    /// <summary>
    /// Command to rename a folder.
    /// </summary>
    [Command(PackageIds.RenameFolder)]
    internal sealed class RenameFolderCommand : BaseCommand<RenameFolderCommand>
    {
        protected override async Task ExecuteAsync(OleMenuCmdEventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var currentItem = FavoritesContextMenuController.CurrentItem;

            if (currentItem is FavoriteFolderNode folderNode)
            {
                var newName = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter new folder name:",
                    "Rename Folder",
                    folderNode.Item.Name);

                if (!string.IsNullOrWhiteSpace(newName) && newName.Trim() != folderNode.Item.Name)
                {
                    FavoritesManager.Instance.RenameFolder(folderNode.Item, newName.Trim());
                }
            }
        }
    }
}
