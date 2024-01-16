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
    public class InvitedToEventBLL:IInvitedToEventBLL
    {
        IInvitedToEventAction _InvitedToEventAction;

        public InvitedToEventBLL(IInvitedToEventAction invitedToEventAction)
        {
            _InvitedToEventAction = invitedToEventAction;
    }


        // מקבל את כל המוזמנים לארוע
        public  List<InvitedToEventDto> getAllInvitedToEventd()
        {
            List<InvitedToEvent> list = _InvitedToEventAction.GetAllInvitedToEventd();
            List<InvitedToEventDto> InvitedToEvents = new List<InvitedToEventDto>();

            foreach (InvitedToEvent item in list)
            {
                InvitedToEventDto i = Converts.convertoDtoInvitedToEvent(item);
                InvitedToEvents.Add(i);
            }
            return InvitedToEvents;
        }
        //הוספת מוזמן לארוע
        public bool addInvitedToEvent(InvitedToEventDto i)
        {
            try
            {
                InvitedToEvent InvitedToEvents = Converts.convertToTblInvitedToEvent(i);
                _InvitedToEventAction.AddInvitedToEvent(InvitedToEvents);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        //מחיקת מוזמן לארוע
        public  List<InvitedToEventDto> deleteInvitedToEventd(int id)
        {
            List<InvitedToEvent> list = _InvitedToEventAction.dellInvitedToEvent(id);
            return getAllInvitedToEventd();
        }
        //עדכון מוזמן לארוע
        public  List<InvitedToEventDto> updateInvitedToEventd(InvitedToEventDto i, int id)
        {
            InvitedToEvent InvitedToEventdToUpdate = Converts.convertToTblInvitedToEvent(i);
            List<InvitedToEvent> newList = _InvitedToEventAction.updateInvitedToEvent(id, InvitedToEventdToUpdate);
            return getAllInvitedToEventd();
        }
        //מביא רשימה לפי  מייל
        public  List<InvitedToEventDto> InvitedToEventbyEmail(string email)
        {
            List<InvitedToEvent> j = _InvitedToEventAction.GetAllInvitedToEventd().Where(n => n.EmailInvited.Equals(email)).ToList();
            if (j == null)
               return null;
            List<InvitedToEventDto> p = Converts.convertoDtoInvitedToEventList(j);
            return p;
        }
    }
}
