[marketplace]: https://marketplace.visualstudio.com/items?itemName=MadsKristensen.SolutionFavorites
[vsixgallery]: http://vsixgallery.com/extension/SolutionFavorites.9f81ec6e-5c91-4809-9dde-9b3166c327fd/
[repo]: https://github.com/madskristensen/SolutionFavorites

# Solution Favorites for Visual Studio

[![Build](https://github.com/madskristensen/SolutionFavorites/actions/workflows/build.yaml/badge.svg)](https://github.com/madskristensen/SolutionFavorites/actions/workflows/build.yaml)

Download this extension from the [Visual Studio Marketplace][marketplace]
or get the [CI build][vsixgallery].

----------------------------------------

Stop hunting through deep folder hierarchies! **Solution Favorites** lets you pin your most-used files to a dedicated **Favorites** node at the top of Solution Explorer for instant access.

![Favorites Node](art/favorites-node.png)

**Organize** with virtual folders | **Rearrange** with drag & drop | **Share** with your team via source control

## Getting Started

1. Right-click any file in Solution Explorer
2. Select **Add to Favorites**
3. Access your pinned files from the **Favorites** node at the top of Solution Explorer

![Add to Favorites](art/add-to-favorites.png)

## Features

### Pin Files for Quick Access

Double-click any favorite to open it instantly. No more navigating through deep folder structures to find your frequently-used files.

### Virtual Folders

Organize your favorites into a custom folder structure:

- Create folders at the root level or nested within other folders
- Rename or remove folders from the right-click context menu
- Build a hierarchy that matches how you think about your project

### Drag and Drop

Rearrange your favorites effortlessly:

- Drag files between folders
- Move files back to the root Favorites node
- Reorganize folders within the hierarchy

### Toggle Visibility

Use the **Toggle Favorites** button on the Solution Explorer toolbar to show or hide the Favorites node when you need more space.

![Toggle visibility](art/toggle-visibility.png)

### Missing File Detection

Files that have been moved or deleted are clearly indicated with:
- A warning icon overlay
- Italic text style
- "File not found" tooltip

### File Actions

Right-click any favorite file for quick actions:

![Context Menu](art/context-menu.png)

- **Remove from Favorites** - Unpin the file
- **Open Containing Folder** - Open in Windows Explorer
- **Copy Full Path** - Copy the path to clipboard

## Team Sharing

Favorites are stored in a `favorites.json` file in the solution directory. Commit this file to source control to share favorites across your team.

**Prefer personal favorites?** Add `favorites.json` to your `.gitignore` file.

## How can I help?

If you enjoy using the extension, please give it a ★★★★★ rating on the [Visual Studio Marketplace][marketplace].

Found a bug or have a feature idea? Head over to the [GitHub repo][repo] to open an issue.

Pull requests are enthusiastically welcomed!

If you find this extension saves you time, please consider [sponsoring me on GitHub](https://github.com/sponsors/madskristensen).