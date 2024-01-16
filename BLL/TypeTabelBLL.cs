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
    public class TypeTabelBLL:ITypeTabelBLL
    {
        ITypeTabelAction _TypeTabelAction;
        public TypeTabelBLL(ITypeTabelAction typeTabelAction)
        {
            _TypeTabelAction = typeTabelAction;
        }

        // מקבל את כל המוזמנים לארוע
        public  List<TypeTabelsDto> getAllTypeTabel()
        {
            List<TypeTabel> list = _TypeTabelAction.GetAllTypeTabel();
            List<TypeTabelsDto> TypeTabels = new List<TypeTabelsDto>();

            foreach (TypeTabel item in list)
            {
                TypeTabelsDto i = Converts.convertoDtoTypeTabels(item);
                TypeTabels.Add(i);
            }
            return TypeTabels;
        }
        //הוספת מוזמן לארוע
        public  List<TypeTabelsDto> addTypeTabel(TypeTabelsDto i)
        {

            TypeTabel TypeTabels = Converts.convertToTblTypeTabelTabel(i);
            List<TypeTabel> list = _TypeTabelAction.AddTypeTabel(TypeTabels);
            return getAllTypeTabel();
        }

        //מחיקת מוזמן לארוע
        public  List<TypeTabelsDto> deleteTypeTabel(int id)
        {
            List<TypeTabel> list = _TypeTabelAction.dellTypeTabel(id);
            return getAllTypeTabel();
        }
        //עדכון מוזמן לארוע
        public  List<TypeTabelsDto> updateTypeTabel(TypeTabelsDto i, int id)
        {
            TypeTabel TypeTabelToUpdate = Converts.convertToTblTypeTabelTabel(i);
            List<TypeTabel> newList = _TypeTabelAction.updateTypeTabel(id, TypeTabelToUpdate);
            return getAllTypeTabel();
        }
    }
}
