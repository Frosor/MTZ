using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MTZ
{
    class MSG
    {
        public MSG()
        {

        }

        public MSG(string messages, string subjects, string sender)
        {
            this.sender = sender;
            this.messages = messages;
            this.subjects = subjects;
        }

        public string messages;
        public string subjects; 
        public string sender;
    }
}
