using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL.Models;
using Microsoft.EntityFrameworkCore;

namespace DAL.action
{
    public class OwnerOfEventAction:IOwnerOfEventAction
    {
        //    EmailOrdersContext myDb = new EmailOrdersContext();
        EmailOrdersContext _db;
        public OwnerOfEventAction(EmailOrdersContext db)
        {
            _db = db;
        }

        public   List<OwnerOfEvent> GetAllOwnerOfEvent()
        {
            
            return  _db.OwnerOfEvents.Include(x=>x.EmailOwnerOfEventNavigation).ToList();

        }
        public   List<OwnerOfEvent> AddOwnerOfEvent(OwnerOfEvent i)
        {
            _db.OwnerOfEvents.Add(i);
            _db.SaveChanges();
            return GetAllOwnerOfEvent();
        }
        public   List<OwnerOfEvent> dellOwnerOfEvent(int id)
        {
            OwnerOfEvent p = _db.OwnerOfEvents.Include(x => x.EmailOwnerOfEventNavigation).FirstOrDefault(n => n.IdEvent == id);
            if (p != null)
                _db.OwnerOfEvents.Remove(p);
            _db.SaveChanges();
            return GetAllOwnerOfEvent();
        }
        public   List<OwnerOfEvent> updateOwnerOfEvent(int id, OwnerOfEvent p)
        {
            OwnerOfEvent pUpdate = _db.OwnerOfEvents.Include(x => x.EmailOwnerOfEventNavigation).FirstOrDefault(n => n.IdEvent == id);
            if (pUpdate != null)
            {
              
                pUpdate.EmailOwnerOfEvent = p.EmailOwnerOfEvent;
                pUpdate.DateOfEvent = p.DateOfEvent;
                pUpdate.AdressOfEvent = p.AdressOfEvent;
              
              
                pUpdate.IdTypeEvent = p.IdTypeEvent;
                pUpdate.NameFileInvitation = p.NameFileInvitation;
            }
            _db.SaveChanges();
            return GetAllOwnerOfEvent();
        }
        //public   List<OwnerOfEvent> getOwnerOfEventByEmail(string email)
        //{
        //     //List<OwnerOfEvent> ownerOfEvent = new List<OwnerOfEvent>();
        //    List<OwnerOfEvent>  ownerOfEvent = _db.OwnerOfEvents.Where(n => n.EmailOwnerOfEvent == email).ToList();
        //    //if(ownerOfEvent != null)
        //        return ownerOfEvent;
        //  //  return null;
        //}
    }
}
