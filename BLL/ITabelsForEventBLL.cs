using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public interface ITabelsForEventBLL
    {
        public  List<TabelsForEventDto> getAllTabelsForEvent();
        public  List<TabelsForEventDto> addTabelsForEvent(TabelsForEventDto i);
        public  List<TabelsForEventDto> deleteTabelsForEvent(int id);
        public  List<TabelsForEventDto> updateTabelsForEvent(TabelsForEventDto i, int id);
    }
}
