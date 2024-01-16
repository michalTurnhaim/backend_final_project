using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL.action
{
    public interface ITypeInviteAction
    {
        public  List<TypeInvite> GetAllTypeInvite();
        public  List<TypeInvite> AddTypeInvite(TypeInvite i);
        public  List<TypeInvite> dellTypeInvite(int id);
        public  List<TypeInvite> updateTypeInvite(int id, TypeInvite p);
    }
}
