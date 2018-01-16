using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace codeFestOwn
{
    class bookMarkNode // class for list
    {
        public string key; // key from excel
        public int page_no; // Page number of the key

        public bookMarkNode(string a, int b) //constructor to initialize each node
        {
            key = a;
            page_no = b;
        }

    }
}
