using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Entity;
using DataAccess;


namespace BusinessLogic
{
    public class UserBusiness
    {
        public User GetUserCode(string userName)
        {
            return new UserData().GetUserCode(userName);
        }
    }
}
