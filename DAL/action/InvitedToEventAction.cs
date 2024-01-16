using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL.Models;
using Microsoft.EntityFrameworkCore;

namespace DAL.action
{
    public class InvitedToEventAction:IInvitedToEventAction
    {
        EmailOrdersContext _db;
        public InvitedToEventAction(EmailOrdersContext db)
        {
            _db = db;
        }

        public  List<InvitedToEvent> GetAllInvitedToEventd()
        {
            
            var a= _db.InvitedToEvents.Include(F => F.IdEventNavigation).Include(k=>k.EmailInvitedNavigation).ToList();
            return a;
        }
        public bool AddInvitedToEvent(InvitedToEvent i)
        {
            try
            {
                _db.InvitedToEvents.Add(i);
                _db.SaveChanges();
                return true;
            }
            catch (Exception ex)
            {
                return false;   
            }
        }

        public  List<InvitedToEvent> dellInvitedToEvent(int id)
        {
            InvitedToEvent p = _db.InvitedToEvents.FirstOrDefault(n => n.IdInvitedToEvent==id);
            if (p != null)
                _db.InvitedToEvents.Remove(p);
            _db.SaveChanges();
            return GetAllInvitedToEventd();
        }
        public  List<InvitedToEvent> updateInvitedToEvent(int id, InvitedToEvent p)
        {
            InvitedToEvent pUpdate = _db.InvitedToEvents.FirstOrDefault(n => n.IdInvitedToEvent == id);
            if (pUpdate != null)
            {
                pUpdate.NumGirls=p.NumGirls;
                pUpdate.NumBoys = p.NumBoys;
                pUpdate.NumTeenageGirls = p.NumTeenageGirls;
                pUpdate.NumteenageBoys = p.NumteenageBoys;
                pUpdate.NumDaughterAdults = p.NumDaughterAdults;
                pUpdate.NumSonAdults=p.NumSonAdults;
                pUpdate.IsCome=p.IsCome;
            }
            _db.SaveChanges();
            return GetAllInvitedToEventd();
        }
    }
}
