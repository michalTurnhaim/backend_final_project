using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using DAL.Models;
using DAL.action;

namespace BLL
{
    public class TypeEventBLL:ITypeEventBLL
    {
        ITypeEventAction _TypeEventAction;
        public TypeEventBLL(ITypeEventAction typeEventAction)
        {
            _TypeEventAction = typeEventAction;
        }

        // מקבל את כל המוזמנים לארוע
        public  List<typeEventDto> getAllTypeEvent()
        {
            List<TypeEvent> list = _TypeEventAction.GetAllTypeEvent();
            List<typeEventDto> TypeEvents = new List<typeEventDto>();

            foreach (TypeEvent item in list)
            {
                typeEventDto t = Converts.convertoDtotypeEventTabel(item);
                TypeEvents.Add(t);
            }
            return TypeEvents;
        }
        //הוספת מוזמן לארוע
        public  List<typeEventDto> addTypeEvent(typeEventDto i)
        {

            TypeEvent TypeEvents = Converts.convertToTbltypeEventTabel(i);
            List<TypeEvent> list = _TypeEventAction.AddTypeEvent(TypeEvents);
            return getAllTypeEvent();
        }

        //מחיקת מוזמן לארוע
        public  List<typeEventDto> deleteTypeEvent(int id)
        {
            List<TypeEvent> list = _TypeEventAction.dellTypeEvent(id);
            return getAllTypeEvent();
        }
        //עדכון מוזמן לארוע
        public  List<typeEventDto> updateTypeEvent(typeEventDto i, int id)
        {
            TypeEvent TypeEventToUpdate = Converts.convertToTbltypeEventTabel(i);
            List<TypeEvent> newList = _TypeEventAction.updateTypeEvent(id, TypeEventToUpdate);
            return getAllTypeEvent();
        }
    }
}
