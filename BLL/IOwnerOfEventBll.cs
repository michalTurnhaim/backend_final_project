using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public interface IOwnerOfEventBll
    {
        public  List<OwnerOfEventDto> getAllOwnerOfEventd();
        public  List<OwnerOfEventDto> addOwnerOfEvent(OwnerOfEventDto i);
        public  List<OwnerOfEventDto> deleteOwnerOfEvent(int id);
        public  List<OwnerOfEventDto> updateOwnerOfEventd(OwnerOfEventDto i, int id);
        //public  List<OwnerOfEventDto> getOwnerOfEventByEmail(string email);
    }
}
