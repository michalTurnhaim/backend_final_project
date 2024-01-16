using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public interface ITabelsForEventAction
    {
        public  List<TabelsForEvent> GetAllTabelsForEvent();
        public  List<TabelsForEvent> AddTabelsForEvent(TabelsForEvent i);
        public  List<TabelsForEvent> dellTabelsForEvent(int id);
        public  List<TabelsForEvent> updateTabelsForEvent(int id, TabelsForEvent t);
    }
}
