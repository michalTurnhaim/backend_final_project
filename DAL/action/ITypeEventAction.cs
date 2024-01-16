using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public interface ITypeEventAction
    {
        public  List<TypeEvent> GetAllTypeEvent();
        public  List<TypeEvent> AddTypeEvent(TypeEvent i);
        public  List<TypeEvent> dellTypeEvent(int id);
        public  List<TypeEvent> updateTypeEvent(int id, TypeEvent t);
    }
}
