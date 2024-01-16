using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public interface ITypeEventBLL
    {
        public  List<typeEventDto> getAllTypeEvent();
        public  List<typeEventDto> addTypeEvent(typeEventDto i);
        public  List<typeEventDto> deleteTypeEvent(int id);
        public  List<typeEventDto> updateTypeEvent(typeEventDto i, int id);
    }
}
