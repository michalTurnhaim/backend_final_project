using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public interface IInvitedAction
    {
        public  List<Invited> updateInvited(string email, Invited p);
        public  List<Invited> AddInvited(Invited i);
        public  List<Invited> dellInvited(string email);
        public  List<Invited> GetAllInvited();
    }
}
