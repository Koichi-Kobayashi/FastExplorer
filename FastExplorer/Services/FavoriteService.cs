using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Cysharp.Text;
using FastExplorer.Models;

namespace FastExplorer.Services
{
    /// <summary>
    /// お気に入りの管理を行うサービス
    /// </summary>
    public class FavoriteService
    {
        private readonly string _favoritesFilePath;
        private List<FavoriteItem> _favorites = new();

        /// <summary>
        /// <see cref="FavoriteService"/>クラスの新しいインスタンスを初期化します
        /// </summary>
        public FavoriteService()
        {
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "FastExplorer");
            Directory.CreateDirectory(appDataPath);
            _favoritesFilePath = Path.Combine(appDataPath, "favorites.json");
            LoadFavorites();
        }

        /// <summary>
        /// すべてのお気に入りを取得します
        /// </summary>
        /// <returns>お気に入りのコレクション</returns>
        public IEnumerable<FavoriteItem> GetFavorites()
        {
            // ToList()を削減：直接IEnumerableを返してメモリ割り当てを削減
            return _favorites;
        }

        /// <summary>
        /// お気に入りを追加します
        /// </summary>
        /// <param name="name">表示名</param>
        /// <param name="path">パス</param>
        public void AddFavorite(string name, string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;

            // 既に同じパスが存在する場合は追加しない（Any()を最適化）
            foreach (var existingFavorite in _favorites)
            {
                if (existingFavorite.Path.Equals(path, StringComparison.OrdinalIgnoreCase))
                    return;
            }

            var favorite = new FavoriteItem
            {
                Name = string.IsNullOrWhiteSpace(name) ? Path.GetFileName(path) ?? path : name,
                Path = path
            };

            _favorites.Add(favorite);
            SaveFavorites();
        }

        /// <summary>
        /// お気に入りを削除します
        /// </summary>
        /// <param name="id">削除するお気に入りのID</param>
        public void RemoveFavorite(string id)
        {
            var favorite = _favorites.FirstOrDefault(f => f.Id == id);
            if (favorite != null)
            {
                _favorites.Remove(favorite);
                SaveFavorites();
            }
        }

        /// <summary>
        /// パスでお気に入りを削除します
        /// </summary>
        /// <param name="path">削除するお気に入りのパス</param>
        public void RemoveFavoriteByPath(string path)
        {
            var favorite = _favorites.FirstOrDefault(f => f.Path.Equals(path, StringComparison.OrdinalIgnoreCase));
            if (favorite != null)
            {
                _favorites.Remove(favorite);
                SaveFavorites();
            }
        }

        /// <summary>
        /// お気に入りを保存します
        /// </summary>
        private void SaveFavorites()
        {
            try
            {
                var json = JsonSerializer.Serialize(_favorites, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_favoritesFilePath, json);
            }
            catch
            {
                // エラーハンドリング
            }
        }

        /// <summary>
        /// お気に入りを読み込みます
        /// </summary>
        private void LoadFavorites()
        {
            try
            {
                if (File.Exists(_favoritesFilePath))
                {
                    var json = File.ReadAllText(_favoritesFilePath);
                    _favorites = JsonSerializer.Deserialize<List<FavoriteItem>>(json) ?? new List<FavoriteItem>();
                }
                else
                {
                    // 初回起動時はデフォルトのお気に入りを追加
                    InitializeDefaultFavorites();
                }
            }
            catch
            {
                _favorites = new List<FavoriteItem>();
                InitializeDefaultFavorites();
            }
        }

        /// <summary>
        /// デフォルトのお気に入りを初期化します
        /// </summary>
        private void InitializeDefaultFavorites()
        {
            _favorites.Clear();

            // 標準的なWindowsフォルダを追加
            AddDefaultFolder("デスクトップ", Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            AddDefaultFolder("ダウンロード", Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\Downloads");
            AddDefaultFolder("ドキュメント", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            AddDefaultFolder("ピクチャ", Environment.GetFolderPath(Environment.SpecialFolder.MyPictures));
            AddDefaultFolder("ミュージック", Environment.GetFolderPath(Environment.SpecialFolder.MyMusic));
            AddDefaultFolder("ビデオ", Environment.GetFolderPath(Environment.SpecialFolder.MyVideos));

            // ごみ箱（特殊なパス）
            AddDefaultFolder("ごみ箱", "shell:RecycleBinFolder");

            // ドライブを追加
            try
            {
                var drives = Directory.GetLogicalDrives();
                foreach (var drive in drives)
                {
                    try
                    {
                        var driveInfo = new DriveInfo(drive);
                        if (driveInfo.IsReady)
                        {
                            // 文字列補間を最適化（ZString.Concatを使用してボクシングを回避）
                            var driveLetter = drive.TrimEnd('\\');
                            var driveName = string.IsNullOrEmpty(driveInfo.VolumeLabel) 
                                ? ZString.Concat("ローカルディスク (", driveLetter, ")")
                                : ZString.Concat(driveInfo.VolumeLabel, " (", driveLetter, ")");
                            AddDefaultFolder(driveName, drive);
                        }
                    }
                    catch
                    {
                        // ドライブにアクセスできない場合はスキップ
                        continue;
                    }
                }
            }
            catch
            {
                // エラーハンドリング
            }

            // OneDriveを追加（存在する場合）
            var oneDrivePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "OneDrive");
            if (Directory.Exists(oneDrivePath))
            {
                AddDefaultFolder("OneDrive", oneDrivePath);
            }

            // ネットワークを追加
            AddDefaultFolder("ネットワーク", "shell:NetworkPlacesFolder");

            SaveFavorites();
        }

        /// <summary>
        /// デフォルトフォルダを追加します（存在する場合のみ）
        /// </summary>
        /// <param name="name">フォルダ名</param>
        /// <param name="path">フォルダパス</param>
        private void AddDefaultFolder(string name, string path)
        {
            // shell:で始まるパスは常に追加（ごみ箱、ネットワークなど）
            if (path.StartsWith("shell:"))
            {
                _favorites.Add(new FavoriteItem
                {
                    Name = name,
                    Path = path
                });
                return;
            }

            // 通常のパスの場合は存在確認
            if (Directory.Exists(path))
            {
                _favorites.Add(new FavoriteItem
                {
                    Name = name,
                    Path = path
                });
            }
        }
    }
}

