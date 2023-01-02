using System;
using System.Drawing;
using System.IO;
using System.Runtime.ConstrainedExecution;
using ClosedXML.Excel;
using ClosedXML.Report.Utils;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelToLightFile;

public class LoadFile
{
    public static int width = 10, height = 50;
    public static byte[] AllBytes;
    public static void Main()
    {
        IXLWorkbook wb = GetWorkBookFromFile("S:/Me/Documents/LightShowV2.xlsx");
        List<IXLWorksheet> worksheets = wb.Worksheets.ToList();
        AllBytes = new byte[((width * height * 3) + 2 ) * (worksheets.Count - 1)];
        int offset = 0;
        foreach (IXLWorksheet sheet in worksheets)
        {
            if (sheet.Name.Equals("TEMPLATE"))
            {
                continue;
            }
            string value = sheet.Cell(1, 1).Value.ToString();
            var bytes = Conversion.ConvertToBytes(GetColorListForSheet(wb, sheet), value);
            for (int i = 0; i < bytes.Length; i++)
            {
                AllBytes[offset+i] = bytes[i];
            }
            offset += (width * height * 3) + 2;
            Console.WriteLine($"New Offset:{offset}");
        }
        File.WriteAllBytes("S:/me/Documents/LightShowV2.bin", AllBytes);
        Console.WriteLine($"Wrote {AllBytes.Length} Bytes to ");
    }

    public static IXLWorkbook GetWorkBookFromFile(string file)
    {
        return new XLWorkbook(file);
    }
    
    public static List<System.Drawing.Color> Load(string file)
    {
        var wb = new XLWorkbook(file);
        var ws = wb.Worksheets.Last();
        var cell = ws.Cell(2, 2);
        var color = cell.Style.Fill.BackgroundColor.Color;

        return GetColorListForSheet(wb,ws);
    }

    private static System.Drawing.Color convertColor(IXLWorkbook wb, IXLCell cell)
    {
        switch (cell.Style.Fill.BackgroundColor.ColorType)
        {
            case XLColorType.Color:
                return cell.Style.Fill.BackgroundColor.Color;
                break;
            case XLColorType.Theme:
                switch (cell.Style.Fill.BackgroundColor.ThemeColor)
                {
                    case XLThemeColor.Accent1:
                        return wb.Theme.Accent1.Color;
                        break;
                    case XLThemeColor.Accent2:
                        return wb.Theme.Accent2.Color;
                        break;
                    case XLThemeColor.Accent3:
                        return wb.Theme.Accent3.Color;
                        break;
                    case XLThemeColor.Accent4:
                        return wb.Theme.Accent4.Color;
                        break;
                    case XLThemeColor.Accent5:
                        return wb.Theme.Accent5.Color;
                        break;
                    case XLThemeColor.Accent6:
                        return wb.Theme.Accent6.Color;
                        break;
                    default:
                        break;
                }
                break;
            default:
                break;
        }
        return System.Drawing.Color.Black;
    }



    private static List<System.Drawing.Color> GetColorListForSheet(IXLWorkbook wb, IXLWorksheet ws)
    {
        List<System.Drawing.Color> colors = new();
        for (int i = 2; i<= width+1; i++)
        {
            if (i % 2 == 0)
            {
                for (int j = height + 1; j > 1; j--)
                {
                    var cell = ws.Cell(j, i);
                    //Console.WriteLine($"Row {j} Col: {i}");
                    colors.Add(convertColor(wb, cell));
                }
            }
            else
            {
                for (int j = 2; j <= height + 1; j++)
                {
                    var cell = ws.Cell(j, i);
                    //Console.WriteLine($"Row {j} Col: {i}");
                    colors.Add(convertColor(wb, cell));
                }
            }
        }
        return colors;
    }
}