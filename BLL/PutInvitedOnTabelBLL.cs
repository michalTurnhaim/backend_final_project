using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL.Models;
using DAL.action;
using DTO;

namespace BLL
{
    public class PutInvitedOnTabelBLL:IPutInvitedOnTabelBLL
    {
        IPutInvitedOnTabelAction _PutInvitedOnTabelAction;
        public PutInvitedOnTabelBLL(IPutInvitedOnTabelAction putInvitedOnTabelAction)
        {
            _PutInvitedOnTabelAction = putInvitedOnTabelAction;
        }
        // מקבל את שיבוץ מוזמנים בשולחנות
        public  List<PutInvitedOnTabelDto> getAllPutInvitedOnTabel()
        {
            List<PutInvitedOnTabel> list = _PutInvitedOnTabelAction.GetAllPutInvitedOnTabel();
            List<PutInvitedOnTabelDto> InvitedToEvents = new List<PutInvitedOnTabelDto>();

            foreach (PutInvitedOnTabel item in list)
            {
                PutInvitedOnTabelDto i = Converts.convertoDtoPutInvitedOnTabel(item);
                InvitedToEvents.Add(i);
            }
            return InvitedToEvents;
        }
        //הוספת שיבוץ מוזמנים בשולחנות
        public  List<PutInvitedOnTabelDto> addPutInvitedOnTabel(PutInvitedOnTabelDto i)
        {

            PutInvitedOnTabel InvitedToEvents = Converts.convertToTblPutInvitedOnTabel(i);
            List<PutInvitedOnTabel> list = _PutInvitedOnTabelAction.AddPutInvitedOnTabel(InvitedToEvents);
            return getAllPutInvitedOnTabel();
        }

        //מחיקת שיבוץ מוזמנים בשולחנות
        public  List<PutInvitedOnTabelDto> deletePutInvitedOnTabel(int id)
        {
            List<PutInvitedOnTabel> list = _PutInvitedOnTabelAction.dellPutInvitedOnTabel(id);
            return getAllPutInvitedOnTabel();
        }
        //עדכון שיבוץ מוזמנים בשולחנות
        public  List<PutInvitedOnTabelDto> updatePutInvitedOnTabel(PutInvitedOnTabelDto i, int id)
        {
            PutInvitedOnTabel InvitedToEventdToUpdate = Converts.convertToTblPutInvitedOnTabel(i);
            List<PutInvitedOnTabel> newList = _PutInvitedOnTabelAction.updatePutInvitedOnTabel(id, InvitedToEventdToUpdate);
            return getAllPutInvitedOnTabel();
        }
    }
}
