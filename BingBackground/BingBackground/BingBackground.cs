using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace BingBackground
{
    class BingBackground
    {
        static void Main(String[] args)
        {
            String urlBase = GetBackgroundURLBase();
            Image background = DownloadBackground(urlBase + GetResolutionExtension(urlBase));
            SaveBackground(background);
            SetBackground(background, PicturePosition.Fill);
        }
        /// <summary>
        /// Downloads the JSON data for the Bing Image of the Day
        /// </summary>
        /// <returns>JSON data for the Bing Image of the Day</returns>
        private static dynamic DownloadJSON()
        {
            using (WebClient webClient = new WebClient())
            {
                Console.WriteLine("Downloading JSON...");
                String jsonString = webClient.DownloadString("https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US");
                return JsonConvert.DeserializeObject<dynamic>(jsonString);
            }
        }
        /// <summary>
        /// Gets the base URL for the Bing Image of the Day
        /// </summary>
        /// <returns>Base URL of the Bing Image of the Day</returns>
        private static String GetBackgroundURLBase()
        {
            dynamic jsonObject = DownloadJSON();
            return "https://www.bing.com" + jsonObject.images[0].urlbase;
        }
        /// <summary>
        /// Gets the title for the Bing Image of the Day
        /// </summary>
        /// <returns>Title of the Bing Image of the Day</returns>
        private static String GetBackgroundTitle()
        {
            dynamic jsonObject = DownloadJSON();
            String copyrightText = jsonObject.images[0].copyright;
            return copyrightText.Substring(0, copyrightText.IndexOf(" ("));
        }
        /// <summary>
        /// Checks to see if website at URL exists
        /// </summary>
        /// <param name="URL">The URL to check for existence</param>
        /// <returns>Whether or not website exists at URL</returns>
        private static Boolean WebsiteExists(String url)
        {
            try
            {
                WebRequest request = WebRequest.Create(url);
                request.Method = "HEAD";
                HttpWebResponse response = (HttpWebResponse) request.GetResponse();
                return (response.StatusCode == HttpStatusCode.OK);
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// Finds the resolution extension for the Bing Image of the Day URL
        /// </summary>
        /// <param name="URL">The base URL</param>
        /// <returns>The resolution extension for the URL</returns>
        private static String GetResolutionExtension(String url)
        {
            Rectangle resolution = Screen.PrimaryScreen.Bounds;
            String widthByHeight = resolution.Width + "x" + resolution.Height;
            String potentialExtension = "_" + widthByHeight + ".jpg";
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
        /// <summary>
        /// Downloads the Bing Image of the Day
        /// </summary>
        /// <param name="URL">The URL of the Bing Image of the Day</param>
        /// <returns>The Bing Image of the Day</returns>
        private static Image DownloadBackground(String url)
        {
            Console.WriteLine("Downloading background...");
            WebRequest request = WebRequest.Create(url);
            WebResponse reponse = request.GetResponse();
            Stream stream = reponse.GetResponseStream();
            return Image.FromStream(stream);
        }
        /// <summary>
        /// Finds the path to My Pictures/Bing Backgrounds/yyyy/M-d-yyyy.bmp
        /// </summary>
        /// <returns>The path to My Pictures/Bing Backgrounds/yyyy/M-d-yyyy.bmp</returns>
        private static String GetBackgroundImagePath()
        {
            String directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Bing Backgrounds", DateTime.Now.Year.ToString());
            Directory.CreateDirectory(directory);
            return Path.Combine(directory, DateTime.Now.ToString("M-d-yyyy") + ".bmp");
        }
        /// <summary>
        /// Saves the Bing Image of the Day to My Pictures/Bing Backgrounds/yyyy/M-d-yyyy.bmp
        /// </summary>
        /// <param name="background">The background image to save</param>
        private static void SaveBackground(Image background)
        {
            Console.WriteLine("Saving background...");
            background.Save(GetBackgroundImagePath(), System.Drawing.Imaging.ImageFormat.Bmp);
        }
        /// <summary>
        /// Different types of PicturePositions to set backgrounds
        /// </summary>
        private enum PicturePosition
        {
            ///<summary>Tiles the picture on the screen</summary>
            Tile,
            ///<summary>Centers the picture on the screen</summary>
            Center,
            ///<summary>Stretches the picture to fit the screen</summary>
            Stretch,
            ///<summary>Fits the picture to the screen</summary>
            Fit,
            ///<summary>Crops the picture to fill the screen</summary>
            Fill
        }
        /// <summary>
        /// Methods that use platform invocation services
        /// </summary>
        internal sealed class NativeMethods
        {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            internal static extern int SystemParametersInfo(int uAction, int uParam, String lpvParam, int fuWinIni);
        }
        /// <summary>
        /// Sets the Bing Image of the Day as the desktop background
        /// </summary>
        /// <param name="background">The background to set</param>
        /// <param name="style">The PicturePosition to use</param>
        private static void SetBackground(Image background, PicturePosition style)
        {
            Console.WriteLine("Setting background...");
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(Path.Combine("Control Panel","Desktop"), true))
            {
                switch (style)
                {
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
            const int SET_DESKTOP_BACKGROUND = 20;
            const int UPDATE_INI_FILE = 1;
            const int SEND_WINDOWS_INI_CHANGE = 2;
            NativeMethods.SystemParametersInfo(SET_DESKTOP_BACKGROUND, 0, GetBackgroundImagePath(), UPDATE_INI_FILE | SEND_WINDOWS_INI_CHANGE);
        }
    }
}