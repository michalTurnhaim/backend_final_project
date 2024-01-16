using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public class TypeInviteAction:ITypeInviteAction
    {
        // static EmailOrdersContext myDb = new EmailOrdersContext();
        EmailOrdersContext _db;
        public TypeInviteAction(EmailOrdersContext db)
        {
            _db = db;
        }

        public  List<TypeInvite> GetAllTypeInvite()
        {
            return _db.TypeInvites.ToList();

        }
        public  List<TypeInvite> AddTypeInvite(TypeInvite i)
        {
            _db.TypeInvites.Add(i);
            _db.SaveChanges();
            return GetAllTypeInvite();
        }
        public  List<TypeInvite> dellTypeInvite(int id)
        {
            TypeInvite p = _db.TypeInvites.FirstOrDefault(n => n.IdTypeInvite == id);
            if (p != null)
                _db.TypeInvites.Remove(p);
            _db.SaveChanges();
            return GetAllTypeInvite();
        }
        public  List<TypeInvite> updateTypeInvite(int id, TypeInvite p)
        {
            TypeInvite pUpdate = _db.TypeInvites.FirstOrDefault(n => n.IdTypeInvite == id);
            if (pUpdate != null)
            {
                pUpdate.IdTypeInvite = p.IdTypeInvite;
                pUpdate.NameTypeInvite = p.NameTypeInvite;
            }
            _db.SaveChanges();
            return GetAllTypeInvite();
        }
    }
}
