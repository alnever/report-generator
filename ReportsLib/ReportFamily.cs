using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLib
{
    public class ReportFamily
    {
        public int ID;
        public int Pos;
        public string Name;
        public string DispName;
        public int Visible;
        public int FrSelect;
        public int RepCount;

        public ReportFamily(int aId, int aPos, string aName, string aDispName, int aVisible, int aFrSelect, int aRepCount = 0)
        {
            ID = aId;
            Pos = aPos;
            Name = aName;
            DispName = aDispName;
            Visible = aVisible;
            FrSelect = aFrSelect;
            RepCount = aRepCount;
        }

    }
}
