using DAL.action;
using DAL.Models;
using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public class TabelsForEventBLL:ITabelsForEventBLL
    {
        ITabelsForEventAction _TabelsForEventAction;
        public TabelsForEventBLL(ITabelsForEventAction tabelsForEventAction)
        {
            _TabelsForEventAction = tabelsForEventAction;
        }

        public  List<TabelsForEventDto> getAllTabelsForEvent()
        {
            List<TabelsForEvent> list = _TabelsForEventAction.GetAllTabelsForEvent();
            List<TabelsForEventDto> TabelsForEvents = new List<TabelsForEventDto>();

            foreach (TabelsForEvent item in list)
            {
                TabelsForEventDto i = Converts.convertoDtotabalesForEventTabel(item);
                TabelsForEvents.Add(i);
            }
            return TabelsForEvents;
        }
        //הוספת מוזמן לארוע
        public  List<TabelsForEventDto> addTabelsForEvent(TabelsForEventDto i)
        {

            TabelsForEvent TabelsForEvents = Converts.convertToTblTabelsForEventTabel(i);
            List<TabelsForEvent> list = _TabelsForEventAction.AddTabelsForEvent(TabelsForEvents);
            return getAllTabelsForEvent();
        }

        //מחיקת מוזמן לארוע
        public  List<TabelsForEventDto> deleteTabelsForEvent(int id)
        {
            List<TabelsForEvent> list = _TabelsForEventAction.dellTabelsForEvent(id);
            return getAllTabelsForEvent();
        }
        //עדכון מוזמן לארוע
        public  List<TabelsForEventDto> updateTabelsForEvent(TabelsForEventDto i, int id)
        {
            TabelsForEvent TabelsForEventdToUpdate = Converts.convertToTblTabelsForEventTabel(i);
            List<TabelsForEvent> newList = _TabelsForEventAction.updateTabelsForEvent(id, TabelsForEventdToUpdate);
            return getAllTabelsForEvent();
        }
    }
}
