using MoreLinq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToLightFile
{
    public static class Conversion
    {
        public static byte[] ConvertToBytes(List<System.Drawing.Color> colors, string frameDuration)
        {
            byte[] bytes = new byte[colors.Count * 3 + 2];
            for(int i = 0; i < colors.Count; i++)
            {
                bytes[i*3] = colors[i].R;
                bytes[i*3 + 1] = colors[i].G;
                bytes[i*3 + 2] = colors[i].B;
                //Console.WriteLine($"{colors[i].R} {colors[i].G} {colors[i].B}");
            }
            var duration = frameDuration.Split('.');
            if(duration.Length < 2)
            {
                duration = duration.Append("0").ToArray();
            }
            _ = int.TryParse(duration[0], out int high);
            _ = int.TryParse(duration[1], out int low);
            bytes[bytes.Length - 2] = (byte)high;
            bytes[bytes.Length - 1] = (byte)low;
            return bytes;
        }
    }
}
