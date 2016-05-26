#region

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Serialization;
using Microsoft.Win32;
using Newtonsoft.Json;

#endregion

namespace BingBackground
{
    internal class BingBackground
    {
        private const string imageDownloadLink = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt={0}";
        private const int SW_HIDE = 0;

        private static void Main()
        {
            Settings settings;
            using (FileStream fileStream = new FileStream(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Settings.xml"), FileMode.Open))
            {
                settings = (Settings)new XmlSerializer(typeof(Settings)).Deserialize(fileStream);
            }

            if (settings.HideWindow)
            {
                ShowWindow(GetConsoleWindow(), SW_HIDE);
            }

            HashSet<string> images = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
            HashSet<string> stringSet = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);

            foreach (Market market in settings.Markets)
            {
                try
                {
                    string backgroundUrlBase = GetBackgroundUrlBase(market);
                    string imageName = Path.GetFileName(backgroundUrlBase).Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries)[0];

                    if (!stringSet.Contains(imageName))
                    {
                        stringSet.Add(imageName);
                        string imageUrl = backgroundUrlBase + GetResolutionExtension(backgroundUrlBase);
                        try
                        {
                            Image background = DownloadBackground(imageUrl);
                            string str = SaveBackground(imageName, background);
                            images.Add(str);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Failed to download file: {0} because of error: {1}.", imageUrl, ex.Message);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Skipped downloading file as it was already downloaded previously.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to download image from market {0} because of error: {1}.", market.MarketName, ex.Message);
                }

                if (settings.DownloadImageFromFirstMarketOnly)
                {
                    break;
                }
            }

            if (settings.DownloadImageFromFirstMarketOnly && images.Count > 0)
            {
                SetBackground(images.First(), PicturePosition.Stretch);
            }
        }

        private static object DownloadJson(Market market)
        {
            using (WebClient webClient = new WebClient())
            {
                Console.WriteLine("Downloading JSON for {0}...", market.MarketName);
                return JsonConvert.DeserializeObject<object>(webClient.DownloadString(string.Format(CultureInfo.InvariantCulture, imageDownloadLink, market.MarketId)));
            }
        }

        private static string GetBackgroundUrlBase(Market market)
        {
            dynamic obj1 = DownloadJson(market);
            return "https://www.bing.com" + obj1.images[0].urlbase;
        }

        private static bool WebsiteExists(string url)
        {
            try
            {
                WebRequest webRequest = WebRequest.Create(url);
                webRequest.Method = "HEAD";
                using (var webResponse = (HttpWebResponse)webRequest.GetResponse())
                {
                    return webResponse.StatusCode == HttpStatusCode.OK;
                }
            }
            catch
            {
                return false;
            }
        }

        private static string GetResolutionExtension(string url)
        {
            Rectangle resolution = Screen.PrimaryScreen.Bounds;
            string widthByHeight = resolution.Width + "x" + resolution.Height;
            string potentialExtension = "_" + widthByHeight + ".jpg";
            if (WebsiteExists(url + potentialExtension))
            {
                Console.WriteLine("Background for " + widthByHeight + " found.");
                return potentialExtension;
            }
            else
            {
                Console.WriteLine("No background for " + widthByHeight + " was found.");
                Console.WriteLine("Using 1920x1080 instead.");
                return "_1920x1080.jpg";
            }
        }

        private static Image DownloadBackground(string url)
        {
            Console.WriteLine("Downloading background {0}...", url);

            WebRequest request = WebRequest.Create(url);

            using (WebResponse response = request.GetResponse())
            {
                using (var stream = response.GetResponseStream())
                {
                    return Image.FromStream(stream);
                }
            }
        }

        private static string GetBackgroundImagePath(string imageName)
        {
            string str = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Bing Backgrounds", DateTime.Now.Year.ToString());
            Directory.CreateDirectory(str);
            return Path.Combine(str, imageName + ".png");
        }

        private static string SaveBackground(string imageName, Image background)
        {
            Console.WriteLine("Saving background...");
            string backgroundImagePath = GetBackgroundImagePath(imageName);
            background.Save(backgroundImagePath, ImageFormat.Png);
            return backgroundImagePath;
        }

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private static void SetBackground(string image, PicturePosition style)
        {
            Console.WriteLine("Setting background...");
            using (RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(Path.Combine("Control Panel", "Desktop"), true))
            {
                switch (style)
                {
                    case PicturePosition.Tile:
                        registryKey.SetValue("PicturePosition", "0");
                        registryKey.SetValue("TileWallpaper", "1");
                        break;
                    case PicturePosition.Center:
                        registryKey.SetValue("PicturePosition", "0");
                        registryKey.SetValue("TileWallpaper", "0");
                        break;
                    case PicturePosition.Stretch:
                        registryKey.SetValue("PicturePosition", "2");
                        registryKey.SetValue("TileWallpaper", "0");
                        break;
                    case PicturePosition.Fit:
                        registryKey.SetValue("PicturePosition", "6");
                        registryKey.SetValue("TileWallpaper", "0");
                        break;
                    case PicturePosition.Fill:
                        registryKey.SetValue("PicturePosition", "10");
                        registryKey.SetValue("TileWallpaper", "0");
                        break;
                }
            }

            const int SetDesktopBackground = 20;
            const int UpdateIniFile = 1;
            const int SendWindowsIniChange = 2;
            NativeMethods.SystemParametersInfo(SetDesktopBackground, 0, image, UpdateIniFile | SendWindowsIniChange);
        }

        private enum PicturePosition
        {
            Tile,
            Center,
            Stretch,
            Fit,
            Fill
        }

        internal sealed class NativeMethods
        {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            internal static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
        }
    }
}