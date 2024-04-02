using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace GraphLibrary
{
    static class VisioColor
    {
        public static string ColorToRgb(string color)
        {
            System.Drawing.Color slateBlue = System.Drawing.Color.FromName(color);
            byte g = slateBlue.G;
            byte b = slateBlue.B;
            byte r = slateBlue.R;
            return $"RGB({g},{b},{r})";
           
        }
    }
}
