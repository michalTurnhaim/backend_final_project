using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public class TabelsForEventAction:ITabelsForEventAction
    {

        //    EmailOrdersContext myDb = new EmailOrdersContext();
        EmailOrdersContext _db;
        public TabelsForEventAction(EmailOrdersContext db)
        {
            _db = db;
        }

        public   List<TabelsForEvent> GetAllTabelsForEvent()
        {
            return _db.TabelsForEvents.ToList();

        }
        public   List<TabelsForEvent> AddTabelsForEvent(TabelsForEvent i)
        {
            _db.TabelsForEvents.Add(i);
            _db.SaveChanges();
            return GetAllTabelsForEvent();
        }
        public   List<TabelsForEvent> dellTabelsForEvent(int id)
        {
            TabelsForEvent p = _db.TabelsForEvents.FirstOrDefault(n => n.IdTabels == id);
            if (p != null)
                _db.TabelsForEvents.Remove(p);
            _db.SaveChanges();
            return GetAllTabelsForEvent();
        }
        public   List<TabelsForEvent> updateTabelsForEvent(int id, TabelsForEvent t)
        {
            TabelsForEvent tUpdate = _db.TabelsForEvents.FirstOrDefault(n => n.IdTabels == id);
            if (tUpdate != null)
            {
                tUpdate.IdTabels = t.IdTabels;
                tUpdate.IdTypeTabels = t.IdTypeTabels;
                tUpdate.IdTypeTabelsNavigation = t.IdTypeTabelsNavigation;
                tUpdate.MenOrWomen = t.MenOrWomen;
                tUpdate.IdEventNavigation = t.IdEventNavigation;
                tUpdate.IdEvent = t.IdEvent;
            }
            _db.SaveChanges();
            return GetAllTabelsForEvent();
        }
    }
}
