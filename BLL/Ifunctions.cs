using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public interface Ifunctions
    {
        public int putNameFileInvitationInData(string fileName);
        public int addOwnerOfEvent(OwnerOfEventNew o);
        public bool READExcel(string fileName);
        public void InsertToInvitedToEvent();
        public string returnNameFile(int id);
        public List<OwnerOfEventDto> returnListOfOwnerByEmail(string email);
        public List<InvitedToEventDto> invitedToEventDtoList(int IdEvent);
        public bool BeforSendingEmail( int idevent);
        public bool SendEmail(string email, int password, string fileOrder);
    }
}
