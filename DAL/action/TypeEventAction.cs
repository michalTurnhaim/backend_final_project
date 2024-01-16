using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL.Models;


namespace DAL.action
{
    public class TypeEventAction:ITypeEventAction
    {

        //   static EmailOrdersContext myDb = new EmailOrdersContext();
        EmailOrdersContext _db;
        public TypeEventAction(EmailOrdersContext db)
        {
            _db = db;
        }

        public  List<TypeEvent> GetAllTypeEvent()
        {
            return _db.TypeEvents.ToList();

        }
        public  List<TypeEvent> AddTypeEvent(TypeEvent i)
        {
            _db.TypeEvents.Add(i);
            _db.SaveChanges();
            return GetAllTypeEvent();
        }
        public  List<TypeEvent> dellTypeEvent(int id)
        {
            TypeEvent p = _db.TypeEvents.FirstOrDefault(n => n.IdTypeEvent == id);
            if (p != null)
                _db.TypeEvents.Remove(p);
            _db.SaveChanges();
            return GetAllTypeEvent();
        }
        public  List<TypeEvent> updateTypeEvent(int id, TypeEvent t)
        {
            TypeEvent tUpdate = _db.TypeEvents.FirstOrDefault(n => n.IdTypeEvent == id);
            if (tUpdate != null)
            {
                tUpdate.IdTypeEvent = t.IdTypeEvent;
                tUpdate.NameTypeEvent = t.NameTypeEvent;
            }
            _db.SaveChanges();
            return GetAllTypeEvent();
        }
    }
}

