using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public interface IOwnerOfEventAction
    {
        public List<OwnerOfEvent> GetAllOwnerOfEvent();
        public List<OwnerOfEvent> AddOwnerOfEvent(OwnerOfEvent i);
        public List<OwnerOfEvent> dellOwnerOfEvent(int id);
        public List<OwnerOfEvent> updateOwnerOfEvent(int id, OwnerOfEvent p);
        //public List<OwnerOfEvent> getOwnerOfEventByEmail(string email);
    }
}
