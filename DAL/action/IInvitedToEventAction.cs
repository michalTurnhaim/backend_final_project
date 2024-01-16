using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public interface IInvitedToEventAction
    {
        public  List<InvitedToEvent> GetAllInvitedToEventd();
        public  bool AddInvitedToEvent(InvitedToEvent i);
        public  List<InvitedToEvent> dellInvitedToEvent(int id);
        public  List<InvitedToEvent> updateInvitedToEvent(int id, InvitedToEvent p);

    }
}
