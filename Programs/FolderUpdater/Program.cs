using SHDocVw;
using Shell32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace FolderUpdater
{
    internal static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            List<IntPtr> explorerWindows = new List<IntPtr>();
            IntPtr currentWindow = FindWindowEx(IntPtr.Zero, IntPtr.Zero, "CabinetWClass", null);

            // 全てのExplorerウインドウを探す
            while (currentWindow != IntPtr.Zero)
            {
                explorerWindows.Add(currentWindow);
                currentWindow = FindWindowEx(IntPtr.Zero, currentWindow, "CabinetWClass", null);
            }

            //一番前のフォルダ名を取得する
            string topLocationName = "";
            foreach (var hWnd in explorerWindows)
            {
                StringBuilder title = new StringBuilder(256);
                GetWindowText(hWnd, title, 256);
                topLocationName = title.ToString();
                break;
            }

            //一番前のフォルダ名と同じInternetExplorer をshellWindowsから探すと、File:/// URLが取得できる
            string locationURL = "";
            Shell shell = new Shell();
            ShellWindows shellWindows = (ShellWindows)shell.Windows();
            foreach (InternetExplorer ie in shellWindows)
            {
                if (ie.LocationName == topLocationName)
                {
                    locationURL = ie.LocationURL;
                    break;
                }
            }

            //locationURLが分かれば、そのフォルダに新規にフォルダを作って、直ぐ削除する
            //新規フォルダ名はランダムなものにする
            if (locationURL != "")
            {
                //Debug.WriteLine(locationURL);
                string directoryPath = Uri.UnescapeDataString(locationURL).Replace("file:///", "");

                // ランダムな文字列を生成する
                string input = Guid.NewGuid().ToString();
                byte[] randamByteArray = Encoding.UTF8.GetBytes(input);

                // SHA256 ハッシュ値を計算する
                byte[] sha256ByteArray;
                using (SHA256 sha256 = SHA256.Create())
                {
                    sha256ByteArray = sha256.ComputeHash(randamByteArray);
                }

                // ハッシュ値を16進数の文字列として表示する
                string newFoldername = BitConverter.ToString(sha256ByteArray).Replace("-", "").ToLowerInvariant();
                newFoldername = newFoldername.Substring(0, 16);

                string newDir = Path.Combine(directoryPath, newFoldername);
                Debug.WriteLine(newDir);
                Directory.CreateDirectory(newDir);
                Thread.Sleep(500);
                Directory.Delete(newDir);
            }
            else
            {
                MessageBox.Show("アクティブなエクスプローラは有りません。", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
    }
}