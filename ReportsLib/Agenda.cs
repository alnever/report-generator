using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportsLib
{
    class Agenda
    {
        private int id;
        public int Id
        {
            get
            {
                return id;
            }
        }
        
        private int aitem;
        public int AItem
        {
            get
            {
                return aitem;
            }
        }

        private string comment;
        public string Comment
        {
            get
            {
                return comment;
            }
        }

        private string fulltext;
        public string FullText
        {
            get
            {
                return fulltext;
            }
        }

        private string number;
        public string Number
        {
            get
            {
                return (id == aitem)?number:"";
            }
        }

        private string info;
        public string Info
        {
            get
            {
                return info;
            }
        }

        public Agenda(int aId, int aAitem, string aComment, string aFullText, string aNumber, string aInfo )
        {
            id = aId;
            aitem = aAitem;
            comment = aComment;
            fulltext = aFullText;
            number = aNumber;
            info = aInfo;
        }

    }
}
