using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public interface IPutInvitedOnTabelAction
    {
        public  List<PutInvitedOnTabel> GetAllPutInvitedOnTabel();
        public  List<PutInvitedOnTabel> AddPutInvitedOnTabel(PutInvitedOnTabel i);
        public  List<PutInvitedOnTabel> dellPutInvitedOnTabel(int id);
        public  List<PutInvitedOnTabel> updatePutInvitedOnTabel(int id, PutInvitedOnTabel p);
    }
}
