using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public interface IInvitedToEventBLL
    {
       
        public  List<InvitedToEventDto> getAllInvitedToEventd();
        public bool addInvitedToEvent(InvitedToEventDto i);
        public  List<InvitedToEventDto> deleteInvitedToEventd(int id);
        public  List<InvitedToEventDto> updateInvitedToEventd(InvitedToEventDto i, int id);
        public  List<InvitedToEventDto> InvitedToEventbyEmail(string email);
    }
}
