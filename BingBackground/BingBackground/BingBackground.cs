using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace BingBackground {

    class BingBackground {

        private static void Main(string[] args) {
            string urlBase = GetBackgroundUrlBase();
            Image background = DownloadBackground(urlBase + GetResolutionExtension(urlBase));
            WriteTitle(background, GetBackgroundTitle());
            SaveBackground(background);
            SetBackground(GetPosition());
        }

        private static dynamic DownloadJson() {
            using (WebClient webClient = new WebClient()) {
                Console.WriteLine("Downloading JSON...");
                string jsonString = webClient.DownloadString("https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-US");
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

        private static void WriteTitle(Image background, string title) {
            
            if (Properties.Settings.Default.TitlePosition.Length == 0)
                return;
            
            string fontName = "Arial";
            if (Properties.Settings.Default.TitleFont.Length > 0)
                fontName = Properties.Settings.Default.TitleFont;

            int size = 12;
            if (Properties.Settings.Default.TitleFontSize > 0)
                size = Properties.Settings.Default.TitleFontSize;

            Graphics graphicImage = Graphics.FromImage(background);
            var font = new Font(fontName, size, FontStyle.Regular, GraphicsUnit.Pixel);
            var stringSize = graphicImage.MeasureString(title, font);
            var offset = graphicImage.MeasureString("W", font);

            int x = 0;
            int y = 0;

            // Position the text, taking into account:
            //  Tthe position of the taskbar
            //  The visible desktop area (Screen.PrimaryScreen.WorkingArea)
            //  If the image is smaller than the visible desktop area in FIT mode
            //  TODO: Coping with odd combinations of image size and stretch/fit modes
            switch (Properties.Settings.Default.TitlePosition.ToUpper()) {
                case "TOPLEFT":
                    x = (int)offset.Width;
                    if (Screen.PrimaryScreen.WorkingArea.Left > 0)
                        x += Screen.PrimaryScreen.WorkingArea.Left;

                    y = (int)offset.Height;
                    if (Screen.PrimaryScreen.WorkingArea.Top > 0)
                        y += Screen.PrimaryScreen.WorkingArea.Top;

                    break;
                case "TOPRIGHT":
                    if (background.Width < Screen.PrimaryScreen.WorkingArea.Width)
                    {
                        // Image is narrower than visible desktop height so right edge of image is visible in fit mode
                        // offset text from edge of image
                        x = background.Width - (int)stringSize.Width;
                    }
                    else
                        // Part of the image can be obscured by the taskbar so offset using the visible desktop area
                        x = Screen.PrimaryScreen.WorkingArea.Right - (int)stringSize.Width;

                    x -= (int)offset.Width;

                    y = (int)offset.Height;
                    if (Screen.PrimaryScreen.WorkingArea.Top > 0)
                        y += Screen.PrimaryScreen.WorkingArea.Top;

                    break;

                case "BOTTOMRIGHT":
                    if (background.Width < Screen.PrimaryScreen.WorkingArea.Width)
                    {
                        // Image is narrower than visible desktop height so right edge of image is visible in fit mode
                        // offset text from edge of image
                        x = background.Width - (int)stringSize.Width;
                    }
                    else
                        // Part of the image can be obscured by the taskbar so offset using the visible desktop area
                        x = Screen.PrimaryScreen.WorkingArea.Right - (int)stringSize.Width;
                    x -= (int)offset.Width;

                    if (background.Height < Screen.PrimaryScreen.WorkingArea.Height)
                    {
                        // Image is smaller than visible desktop area so offset text from base of image
                        y = background.Height - (int)stringSize.Height;
                    }
                    else
                        y = Screen.PrimaryScreen.WorkingArea.Bottom - (int)stringSize.Height;

                    
                    break;

                case "BOTTOMLEFT":
                default:
                    x = (int)offset.Width;
                    if (Screen.PrimaryScreen.WorkingArea.Left > 0)
                        x += Screen.PrimaryScreen.WorkingArea.Left;

                    if (background.Height < Screen.PrimaryScreen.WorkingArea.Height) {
                        // Image is shorter than visible desktop height so bottom edge of image is visible in fit mode
                        // offset text from edge of image
                        y = background.Height - (int)stringSize.Height;
                    }
                    else
                        // Part of the image can be obscured by the taskbar so offset using the visible desktop area
                        y = Screen.PrimaryScreen.WorkingArea.Bottom - (int)stringSize.Height;
                    break;
            }

            // Set the text colour to black or white, depending on the background for where the text is overlaying
            Rectangle textArea = new Rectangle(x, y, (int)stringSize.Width, (int)stringSize.Height);
            Color textColour = ContrastingColor(GetDominantColour(background, textArea));
            SolidBrush drawBrush = new SolidBrush(textColour);
            
            graphicImage.SmoothingMode = SmoothingMode.AntiAlias;
            
            graphicImage.DrawString(title, font, drawBrush, new Point(x, y));
            //graphicImage.DrawString(title, font, SystemBrushes.WindowText, new Point(x, y));
            graphicImage.Dispose();
        }


        /// <summary>
        /// Find the dominant colour of an area in an image
        /// </summary>
        /// <param name="image">Image to inspec</param>
        /// <param name="area">Portion of the image to restrict the analysis to</param>
        /// <returns>Color representing the average colour of the pixels scanned</returns>
        private static Color GetDominantColour(Image image, Rectangle area) {
            int r = 0;
            int g = 0;
            int b = 0;

            var bmp = new Bitmap(image);

            for (int x = area.Left; x < area.Right; x++) {
                for (int y = area.Top; y < area.Bottom; y++)
                {
                    Color clr = bmp.GetPixel(x, y);
                    r += clr.R;
                    g += clr.G;
                    b += clr.B;
                }
            }

            int total = area.Width * area.Height;
            r /= total;
            g /= total;
            b /= total;

            return Color.FromArgb(r, g, b);
        
        }

        /// <summary>
        /// Calculate if it
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        private static Color ContrastingColor(Color color) {
            // Calculate the perceptive luminance
            // The human eye favors green color... 
            // From http://stackoverflow.com/questions/1855884/determine-font-color-based-on-background-color
            double perceptiveLuminance = 1 - (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255;

            if (perceptiveLuminance < 0.5)
                return Color.Black;
            else
                return Color.White;
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
            string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures), "Bing Backgrounds", DateTime.Now.Year.ToString());
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
            switch (Properties.Settings.Default.Position.ToUpper()) {
                case "TILE":
                    position = PicturePosition.Tile;
                    break;
                case "CENTER":
                    position = PicturePosition.Center;
                    break;
                case "STRETCH":
                    position = PicturePosition.Stretch;
                    break;
                case "FIT":
                    position = PicturePosition.Fit;
                    break;
                case "FILL":
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