using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public interface ITypeTabelAction
    {
        public  List<TypeTabel> GetAllTypeTabel();
        public  List<TypeTabel> AddTypeTabel(TypeTabel i);
        public  List<TypeTabel> dellTypeTabel(int id);
        public  List<TypeTabel> updateTypeTabel(int id, TypeTabel p);
    }
}
