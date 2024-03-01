using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TEP
{
    internal class Data
    {
        public Element element { get; set; }
        public string function { get; set; }
        public string name { get; set; }
        public string number_apart { get; set; }
        public string number_room { get; set; }
        public string floor { get; set; }
        public string section { get; set; } 
        public double area { get; set; }
        public double height { get; set; }
    }
}
