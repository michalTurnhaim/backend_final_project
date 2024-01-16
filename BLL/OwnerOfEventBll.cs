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
    public class OwnerOfEventBll:IOwnerOfEventBll
    {
        // מקבל את כל המוזמנים לארוע
        IOwnerOfEventAction _ownerOfEventAction;
        public OwnerOfEventBll(IOwnerOfEventAction ownerOfEventAction)
        {
            _ownerOfEventAction = ownerOfEventAction;
        }

            public  List<OwnerOfEventDto> getAllOwnerOfEventd()
            {
                List<OwnerOfEvent> list = _ownerOfEventAction.GetAllOwnerOfEvent();
                List<OwnerOfEventDto> OwnerOfEvents = new List<OwnerOfEventDto>();

                foreach (OwnerOfEvent item in list)
                {
                OwnerOfEventDto i = Converts.convertoDtoOwnerOfEvent(item);
                OwnerOfEvents.Add(i);
                }
                return OwnerOfEvents;
            }
            //הוספת מוזמן לארוע
            public List<OwnerOfEventDto> addOwnerOfEvent(OwnerOfEventDto i)
            {

            OwnerOfEvent OwnerOfEvents = Converts.convertToTblOwnerOfEvent(i);
                List<OwnerOfEvent> list = _ownerOfEventAction.AddOwnerOfEvent(OwnerOfEvents);
                return getAllOwnerOfEventd();
            }

            //מחיקת מוזמן לארוע
            public List<OwnerOfEventDto> deleteOwnerOfEvent(int id)
            {
                List<OwnerOfEvent> list = _ownerOfEventAction.dellOwnerOfEvent(id);
                return getAllOwnerOfEventd();
            }
            //עדכון מוזמן לארוע
            public List<OwnerOfEventDto> updateOwnerOfEventd(OwnerOfEventDto i, int id)
            {
            OwnerOfEvent OwnerOfEventdToUpdate = Converts.convertToTblOwnerOfEvent(i);
                List<OwnerOfEvent> newList = _ownerOfEventAction.updateOwnerOfEvent(id, OwnerOfEventdToUpdate);
                return getAllOwnerOfEventd();
            }
            //קבלת בעל ארוע לפי כתובת מייל
            //public List<OwnerOfEventDto> getOwnerOfEventByEmail(string email)
            //{
            //return Converts.convertoDtoOwnerOfEventList(_ownerOfEventAction.getOwnerOfEventByEmail(email));
            //}
        }

    }

