using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL.Models;
namespace DAL.action
{
    public class PutInvitedOnTabelAction:IPutInvitedOnTabelAction
    {
        EmailOrdersContext _db;
        public PutInvitedOnTabelAction(EmailOrdersContext db)
        {
            _db = db;
        }

        public   List<PutInvitedOnTabel> GetAllPutInvitedOnTabel()
        {
            return _db.PutInvitedOnTabels.ToList();

        }
        public   List<PutInvitedOnTabel> AddPutInvitedOnTabel(PutInvitedOnTabel i)
        {
            _db.PutInvitedOnTabels.Add(i);
            _db.SaveChanges();
            return GetAllPutInvitedOnTabel();
        }
        public   List<PutInvitedOnTabel> dellPutInvitedOnTabel(int id)
        {
            PutInvitedOnTabel p = _db.PutInvitedOnTabels.FirstOrDefault(n => n.IdPutInvitedOnTabels == id);
            if (p != null)
                _db.PutInvitedOnTabels.Remove(p);
            _db.SaveChanges();
            return GetAllPutInvitedOnTabel();
        }
        public   List<PutInvitedOnTabel> updatePutInvitedOnTabel(int id, PutInvitedOnTabel p)
        {
            PutInvitedOnTabel pUpdate = _db.PutInvitedOnTabels.FirstOrDefault(n => n.IdPutInvitedOnTabels == id);
            if (pUpdate != null)
            {
                pUpdate.IdInvitedToEvent = p.IdInvitedToEvent;
                pUpdate.IdTabels = p.IdTabels;
            }
            _db.SaveChanges();
            return GetAllPutInvitedOnTabel();
        }
    }
}
