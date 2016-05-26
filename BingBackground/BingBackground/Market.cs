using System.Xml.Serialization;

namespace BingBackground
{
  public class Market
  {
    [XmlAttribute]
    public string MarketId { get; set; }

    [XmlAttribute]
    public string MarketName { get; set; }

    public Market()
    {
    }

    public Market(string marketId, string marketName)
    {
      this.MarketId = marketId;
      this.MarketName = marketName;
    }
  }
}
