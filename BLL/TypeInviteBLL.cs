using DAL.action;
using DAL.Models;
using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public class TypeInviteBLL:ITypeInviteBLL
    {
        ITypeInviteAction _TypeInviteAction;
        public TypeInviteBLL(ITypeInviteAction _TypeInviteAction)
        {
            _TypeInviteAction = _TypeInviteAction;
        }
        // מקבל את כל המוזמנים לארוע
        public  List<TypeInviteDto> getAllTypeInviteBll()
        {
            List<TypeInvite> list = _TypeInviteAction.GetAllTypeInvite();
            List<TypeInviteDto> TypeEvents = new List<TypeInviteDto>();

            foreach (TypeInvite item in list)
            {
                TypeInviteDto t = Converts.convertoDtoTypeInviteTabel(item);
                TypeEvents.Add(t);
            }
            return TypeEvents;
        }
        //הוספת מוזמן לארוע
        public  List<TypeInviteDto> addTypeInvite(TypeInviteDto i)
        {

            TypeInvite TypeInvites = Converts.convertToTblTypeInviteTabel(i);
            List<TypeInvite> list = _TypeInviteAction.AddTypeInvite(TypeInvites);
            return getAllTypeInviteBll();
        }

        //מחיקת מוזמן לארוע
        public  List<TypeInviteDto> deleteTypeInvite(int id)
        {
            List<TypeInvite> list = _TypeInviteAction.dellTypeInvite(id);
            return getAllTypeInviteBll();
        }
        //עדכון מוזמן לארוע
        public  List<TypeInviteDto> updateTypeInvite(TypeInviteDto i, int id)
        {
            TypeInvite TypeInviteToUpdate = Converts.convertToTblTypeInviteTabel(i);
            List<TypeInvite> newList = _TypeInviteAction.updateTypeInvite(id, TypeInviteToUpdate);
            return getAllTypeInviteBll();
        }
    }
}
