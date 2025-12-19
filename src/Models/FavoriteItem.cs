using System.Collections.Generic;
using Newtonsoft.Json;

namespace SolutionFavorites.Models
{
    /// <summary>
    /// Represents a favorite item (file or folder) in a hierarchical structure.
    /// </summary>
    public class FavoriteItem
    {
        /// <summary>
        /// Display name shown in Solution Explorer.
        /// </summary>
        [JsonProperty("name")]
        public string Name { get; set; }

        /// <summary>
        /// Relative path to the file (only for file items, null for folders).
        /// </summary>
        [JsonProperty("path", NullValueHandling = NullValueHandling.Ignore)]
        public string Path { get; set; }

        /// <summary>
        /// Child items (only for folders).
        /// </summary>
        [JsonProperty("children", NullValueHandling = NullValueHandling.Ignore)]
        public List<FavoriteItem> Children { get; set; }

        /// <summary>
        /// Returns true if this is a folder (has children list).
        /// </summary>
        [JsonIgnore]
        public bool IsFolder => Children != null;

        /// <summary>
        /// Creates a new file favorite.
        /// </summary>
        public static FavoriteItem CreateFile(string path, string name = null)
        {
            return new FavoriteItem
            {
                Name = name ?? System.IO.Path.GetFileName(path),
                Path = path
            };
        }

        /// <summary>
        /// Creates a new folder favorite.
        /// </summary>
        public static FavoriteItem CreateFolder(string name)
        {
            return new FavoriteItem
            {
                Name = name,
                Children = new List<FavoriteItem>()
            };
        }
    }

    /// <summary>
    /// Root container for all favorites data.
    /// </summary>
    public class FavoritesData
    {
        [JsonProperty("version")]
        public int Version { get; set; } = 2;

        [JsonProperty("items")]
        public List<FavoriteItem> Items { get; set; } = new List<FavoriteItem>();
    }
}
