using System.Collections.Generic;

namespace BingBackground
{
  public class Settings
  {
    public bool HideWindow { get; set; }

    public bool DownloadImageFromFirstMarketOnly { get; set; }

    public List<Market> Markets { get; set; }

    public Settings()
    {
      this.HideWindow = false;
      this.DownloadImageFromFirstMarketOnly = true;
      this.Markets = new List<Market>();
    }
  }
}
