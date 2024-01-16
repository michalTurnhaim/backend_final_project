using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public interface IPutInvitedOnTabelBLL
    {
        public  List<PutInvitedOnTabelDto> getAllPutInvitedOnTabel();
        public  List<PutInvitedOnTabelDto> addPutInvitedOnTabel(PutInvitedOnTabelDto i);
        public  List<PutInvitedOnTabelDto> deletePutInvitedOnTabel(int id);
        public  List<PutInvitedOnTabelDto> updatePutInvitedOnTabel(PutInvitedOnTabelDto i, int id);
    }
}
