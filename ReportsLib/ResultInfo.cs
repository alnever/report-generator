using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLib
{
    public class ResultInfo
    {
        public int IdRes;
        public int RType;
        public int RSubtype;

        public ResultInfo(int aId, int aType, int aSubtype)
        {
            IdRes = aId;
            RType = aType;
            RSubtype = aSubtype;
        }
    }
}
