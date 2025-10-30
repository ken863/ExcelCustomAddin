using System;
using System.Drawing;

namespace ExcelCustomAddin
{
  /// <summary>
  /// Shared POCO to represent sheet information displayed in the ActionPanel list
  /// </summary>
  public class SheetInfo
  {
    public string Name { get; set; }
    public Color TabColor { get; set; }
    public bool HasTabColor { get; set; }
    public bool IsPinned { get; set; } = false;

    public override string ToString()
    {
      return Name;
    }
  }
}
