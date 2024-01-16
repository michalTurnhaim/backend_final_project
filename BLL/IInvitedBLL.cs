using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public interface IInvitedBLL
    {
        public  List<InvitedDto> getAllInvited();
        public  List<InvitedDto> addInvited(InvitedDto i);
        public  List<InvitedDto> deleteInvited(string id);
        public  List<InvitedDto> updateInvited(InvitedDto i, string id);
        public  InvitedDto login(string email,int password);
        public InvitedDto checEmailIfExists(string email);
    }
}
