using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace BingBackground {

    class BingBackground {

        private static void Main(string[] args) {
            string urlBase = GetBackgroundUrlBase();
            Image background = DownloadBackground(urlBase + GetResolutionExtension(urlBase));
            SaveBackground(background);
            SetBackground(GetPosition());
        }

        private static dynamic DownloadJson() {
            using (WebClient webClient = new WebClient()) {
                Console.WriteLine("Downloading JSON...");
                webClient.Encoding = System.Text.Encoding.UTF8;
                string jsonString = webClient.DownloadString("https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-UK");
                return JsonConvert.DeserializeObject<dynamic>(jsonString);
            }
        }

        private static string GetBackgroundUrlBase() {
            dynamic jsonObject = DownloadJson();
            return "https://www.bing.com" + jsonObject.images[0].urlbase;
        }

        private static string GetBackgroundTitle() {
            dynamic jsonObject = DownloadJson();
            string copyrightText = jsonObject.images[0].copyright;
            return copyrightText.Substring(0, copyrightText.IndexOf(" ("));
        }

        private static bool WebsiteExists(string url) {
            try {
                WebRequest request = WebRequest.Create(url);
                request.Method = "HEAD";
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                return response.StatusCode == HttpStatusCode.OK;
            } catch {
                return false;
            }
        }

        private static string GetResolutionExtension(string url) {
            Rectangle resolution = Screen.PrimaryScreen.Bounds;
            string widthByHeight = resolution.Width + "x" + resolution.Height;
            string potentialExtension = "_" + widthByHeight + ".jpg";
            if (WebsiteExists(url + potentialExtension)) {
                Console.WriteLine("Background for " + widthByHeight + " found.");
                return potentialExtension;
            } else {
                Console.WriteLine("No background for " + widthByHeight + " was found.");
                Console.WriteLine("Using 1920x1080 instead.");
                return "_1920x1080.jpg";
            }
        }

        private static void SetProxy() {
            string proxyUrl = Properties.Settings.Default.Proxy;
            if (proxyUrl.Length > 0) {
                var webProxy = new WebProxy(proxyUrl, true);
                webProxy.Credentials = CredentialCache.DefaultCredentials;
                WebRequest.DefaultWebProxy = webProxy;
            }
        }

        private static Image DownloadBackground(string url) {
            Console.WriteLine("Downloading background...");
            SetProxy();
            WebRequest request = WebRequest.Create(url);
            WebResponse reponse = request.GetResponse();
            Stream stream = reponse.GetResponseStream();
            return Image.FromStream(stream);
        }

        private static string GetBackgroundImagePath() {
            //string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Bing Backgrounds", DateTime.Now.Year.ToString());
            string directory = Path.Combine("C:\\Users\\Bill - Surface Book\\OneDrive - student.xjtlu.edu.cn\\Pictures", "Bing Backgrounds", DateTime.Now.Year.ToString());
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, DateTime.Now.ToString("M-d-yyyy") + ".bmp");
        }

        private static void SaveBackground(Image background) {
            Console.WriteLine("Saving background...");
            background.Save(GetBackgroundImagePath(), System.Drawing.Imaging.ImageFormat.Bmp);
        }

        private enum PicturePosition {
            Tile,
            Center,
            Stretch,
            Fit,
            Fill
        }

        private static PicturePosition GetPosition() {
            PicturePosition position = PicturePosition.Fit;
            switch (Properties.Settings.Default.Position) {
                case "Tile":
                    position = PicturePosition.Tile;
                    break;
                case "Center":
                    position = PicturePosition.Center;
                    break;
                case "Stretch":
                    position = PicturePosition.Stretch;
                    break;
                case "Fit":
                    position = PicturePosition.Fit;
                    break;
                case "Fill":
                    position = PicturePosition.Fill;
                    break;
            }
            return position;
        }

        internal sealed class NativeMethods {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            internal static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
        }

        private static void SetBackground(PicturePosition style) {
            Console.WriteLine("Setting background...");
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(Path.Combine("Control Panel", "Desktop"), true)) {
                switch (style) {
                    case PicturePosition.Tile:
                        key.SetValue("PicturePosition", "0");
                        key.SetValue("TileWallpaper", "1");
                        break;
                    case PicturePosition.Center:
                        key.SetValue("PicturePosition", "0");
                        key.SetValue("TileWallpaper", "0");
                        break;
                    case PicturePosition.Stretch:
                        key.SetValue("PicturePosition", "2");
                        key.SetValue("TileWallpaper", "0");
                        break;
                    case PicturePosition.Fit:
                        key.SetValue("PicturePosition", "6");
                        key.SetValue("TileWallpaper", "0");
                        break;
                    case PicturePosition.Fill:
                        key.SetValue("PicturePosition", "10");
                        key.SetValue("TileWallpaper", "0");
                        break;
                }
            }
            const int SetDesktopBackground = 20;
            const int UpdateIniFile = 1;
            const int SendWindowsIniChange = 2;
            NativeMethods.SystemParametersInfo(SetDesktopBackground, 0, GetBackgroundImagePath(), UpdateIniFile | SendWindowsIniChange);
        }

    }

}