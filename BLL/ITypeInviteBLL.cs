using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public interface ITypeInviteBLL
    {

        public  List<TypeInviteDto> getAllTypeInviteBll();
        public  List<TypeInviteDto> addTypeInvite(TypeInviteDto i);
        public  List<TypeInviteDto> deleteTypeInvite(int id);
        public  List<TypeInviteDto> updateTypeInvite(TypeInviteDto i, int id);
    }
}
