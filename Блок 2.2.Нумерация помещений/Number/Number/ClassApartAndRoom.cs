using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Number
{
    internal class ClassApartAndRoom
    {
        public Group group { get; set; }
        public String ADSK_Номер_квартиры { get; set; }
        public String PNR_Номер_помещения { get; set; }
        public String name { get; set; }
        public ClassApartAndRoom(string aDSK_Номер_квартиры, string pNR_Номер_помещения, string name)
        {
            ADSK_Номер_квартиры = aDSK_Номер_квартиры;
            PNR_Номер_помещения = pNR_Номер_помещения;
            this.name = name;   
        }

        public override bool Equals(object obj)
        {
            if (obj is ClassApartAndRoom group) return name == group.name;
            return false;
        }

    }
}
