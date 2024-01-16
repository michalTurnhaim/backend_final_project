using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public class TypeTabelAction:ITypeTabelAction
    {
        //  EmailOrdersContext myDb = new EmailOrdersContext();
        EmailOrdersContext _db;

        public  List<TypeTabel> GetAllTypeTabel()
        {
            return _db.TypeTabels.ToList();

        }
        public  List<TypeTabel> AddTypeTabel(TypeTabel i)
        {
            _db.TypeTabels.Add(i);
            _db.SaveChanges();
            return GetAllTypeTabel();
        }
        public  List<TypeTabel> dellTypeTabel(int id)
        {
            TypeTabel p = _db.TypeTabels.FirstOrDefault(n => n.IdTypeTabels == id);
            if (p != null)
                _db.TypeTabels.Remove(p);
            _db.SaveChanges();
            return GetAllTypeTabel();
        }
        public  List<TypeTabel> updateTypeTabel(int id, TypeTabel p)
        {
            TypeTabel pUpdate = _db.TypeTabels.FirstOrDefault(n => n.IdTypeTabels == id);
            if (pUpdate != null)
            {
                pUpdate.IdTypeTabels = p.IdTypeTabels;
                pUpdate.NameTypeTabels = p.NameTypeTabels;
            }
            _db.SaveChanges();
            return GetAllTypeTabel();
        }
    }
}
