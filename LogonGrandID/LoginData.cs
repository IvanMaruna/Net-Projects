using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LogonGrandID
{
    public class LoginData
    {
        public string sessionId { get; set; }
        public string redirectUrl { get; set; }
        public string username { get; set; }
        public UserAttributes userAttributes { get; set; }
    }

    public class UserAttributes
    {
        public string givenname { get; set; }
        public string contactid { get; set; }
        public string adress { get; set; }
    }
}
