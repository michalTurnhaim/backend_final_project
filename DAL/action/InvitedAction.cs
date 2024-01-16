using DAL.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public class InvitedAction:IInvitedAction
    {
        // static EmailOrdersContext myDb = new EmailOrdersContext();
        EmailOrdersContext _db;
        public InvitedAction(EmailOrdersContext db)
        {
            _db = db;
        }

        public  List<Invited> GetAllInvited()
        {
            return _db.Inviteds.ToList();

        }
        public  List<Invited> AddInvited(Invited i)
        {
            _db.Inviteds.Add(i);
            _db.SaveChanges();
            return GetAllInvited();
        }
        public  List<Invited> dellInvited(string email)
        {
            Invited p = _db.Inviteds.FirstOrDefault(n => n.EmailInvited.Equals(email));
            if (p != null)
                _db.Inviteds.Remove(p);
            _db.SaveChanges();
            return GetAllInvited();
        }
       
        public  List<Invited> updateInvited(string email, Invited p)
        {
            Invited pUpdate = _db.Inviteds.FirstOrDefault(n => n.EmailInvited.Equals(email));
            if (pUpdate != null)
            {
                pUpdate.FirstNameInvited=p.FirstNameInvited;
                pUpdate.LastNameInvited=p.LastNameInvited;
                pUpdate.PassWord=p.PassWord;
            }


            _db.SaveChanges();
            return GetAllInvited();
        }

    }
}
