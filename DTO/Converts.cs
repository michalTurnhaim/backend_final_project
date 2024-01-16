using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL.Models;


namespace DTO
{
    public class Converts
    {
        //Invited
        // DTOפונקציה שממירה מ טמבל☺ ל 
        public static InvitedDto convertoDtoInvited(Invited i)
        {
            InvitedDto obj = new InvitedDto();
            obj.EmailInvitedDto=i.EmailInvited;
            obj.FirstNameInvitedDto = i.FirstNameInvited;
            obj.LastNameInvitedDto = i.LastNameInvited;
            obj.PassWordDto = i.PassWord;
            return obj;
        }

        // ☺לטמבל DTOפונקציה שממירה מ  
        public static Invited convertTotblInvuted(InvitedDto i)
        {
            Invited obj = new Invited();
            obj.EmailInvited = i.EmailInvitedDto;
            obj.FirstNameInvited = i.FirstNameInvitedDto;
            obj.LastNameInvited = i.LastNameInvitedDto;
            obj.PassWord = i.PassWordDto;
            return obj;
        }
        //DTO פונקציה הממירה מרשימה ל
        public static List<InvitedDto> convertoDtoInvitedList(List<Invited> p)
        {
            List<InvitedDto> obj = new List<InvitedDto>();
            foreach (Invited item in p)
            {
                obj.Add(convertoDtoInvited(item));
            }
            return obj;
        }
        public static List<Invited> convertToTblInvitedList(List<InvitedDto> p)
        {
            List<Invited> productTbls = new List<Invited>();
            foreach (InvitedDto item in p)
            {
                productTbls.Add(convertTotblInvuted(item));
            }
            return productTbls;
        }
        //--------------
        //InvitedToEvent
        //--------------
        // DTOפונקציה שממירה מ טמבל☺ ל 
        public static InvitedToEventDto convertoDtoInvitedToEvent(InvitedToEvent i)
        {
            InvitedToEventDto obj = new InvitedToEventDto();
            
            obj.IdInvitedToEventDto = i.IdInvitedToEvent;
            obj.IdEventDto=i.IdEvent;
            obj.IdTypeInviteDto=i.IdTypeInvite;
            obj.EmailInvitedDto=i.EmailInvited;
            if(i.IdEventNavigation != null)
            {
                if (i.IdEventNavigation.EmailOwnerOfEventNavigation != null) { 
                obj.firstNameOwner = i.IdEventNavigation.EmailOwnerOfEventNavigation.FirstNameInvited;
                obj.lastNameOwner = i.IdEventNavigation.EmailOwnerOfEventNavigation.LastNameInvited;}
            obj.pic = i.IdEventNavigation.NameFileInvitation;
            obj.DateOfEventDto = i.IdEventNavigation.DateOfEvent;

            }
            if (i.EmailInvitedNavigation != null) { 
            obj.LName = i.EmailInvitedNavigation.LastNameInvited;
            obj.FName=i.EmailInvitedNavigation.FirstNameInvited;}
            obj.NumBoysDto=i.NumBoys;
            obj.NumGirlsDto = i.NumGirls;

            obj.NumteenageBoysDto = i.NumteenageBoys;
            obj.NumTeenageGirlsDto = i.NumTeenageGirls;
            obj.NumDaughterAdultsDto=i.NumDaughterAdults;

            obj.NumSonAdultsDto=i.NumSonAdults;

            obj.IsComeDto = i.IsCome;

           

            return obj;
        }

       

        // ☺לטמבל DTOפונקציה שממירה מ  
        public static InvitedToEvent convertToTblInvitedToEvent(InvitedToEventDto i)
        {
            InvitedToEvent obj = new InvitedToEvent();
            obj.IdInvitedToEvent = i.IdInvitedToEventDto;
            obj.IdEvent = i.IdEventDto;
            obj.EmailInvited = i.EmailInvitedDto;


            obj.IdTypeInvite = i.IdTypeInviteDto;

            obj.NumBoys = i.NumBoysDto;
            obj.NumGirls = i.NumGirlsDto;
            obj.IsCome = i.IsComeDto;
            obj.NumteenageBoys = i.NumteenageBoysDto;
            obj.NumTeenageGirls = i.NumTeenageGirlsDto;

            obj.NumDaughterAdults = i.NumDaughterAdultsDto;
            obj.NumSonAdults= i.NumSonAdultsDto ;
            return obj;
        }
        //DTO פונקציה הממירה מרשימה ל
        public static List<InvitedToEventDto> convertoDtoInvitedToEventList(List<InvitedToEvent> p)
        {
            List<InvitedToEventDto> obj = new List<InvitedToEventDto>();
            foreach (InvitedToEvent item in p)
            {
                obj.Add(convertoDtoInvitedToEvent(item));
            }
            return obj;
        }
        public static List<InvitedToEvent> convertToTblInvitedToEventList(List<InvitedToEventDto> p)
        {
            List<InvitedToEvent> productTbls = new List<InvitedToEvent>();
            foreach (InvitedToEventDto item in p)
            {
                productTbls.Add(convertToTblInvitedToEvent(item));
            }
            return productTbls;
        }
        //OwnerOfEvent
        // DTOפונקציה שממירה מ טמבל☺ ל 
        public static OwnerOfEventDto convertoDtoOwnerOfEvent(OwnerOfEvent o)
        {
            if(o!=null)
            {
                OwnerOfEventDto obj = new OwnerOfEventDto();
                obj.AdressOfEventDto = o.AdressOfEvent;
                obj.DateOfEventDto = o.DateOfEvent;
                obj.IdEventDto = o.IdEvent;
                obj.EmailOwnerOfEventDto = o.EmailOwnerOfEvent;
                obj.FirstNameOwnerOfEventDto = o.EmailOwnerOfEventNavigation.FirstNameInvited;
                obj.LastNameOwnerOfEventDto = o.EmailOwnerOfEventNavigation.LastNameInvited;
                obj.NameFileInvitationDto = o.NameFileInvitation;
                return obj;
            }
            return null;
           
        }

        // ☺לטמבל DTOפונקציה שממירה מ  
        public static OwnerOfEvent convertToTblOwnerOfEvent(OwnerOfEventDto i)
        {
            OwnerOfEvent obj = new OwnerOfEvent();
          
            obj.AdressOfEvent = i.AdressOfEventDto;
            obj.DateOfEvent = i.DateOfEventDto;
            obj.IdEvent = i.IdEventDto;
            obj.EmailOwnerOfEvent = i.EmailOwnerOfEventDto;
            obj.IdTypeEvent = i.IdTypeEventDto;
            obj.NameFileInvitation = i.NameFileInvitationDto;
            
            return obj;
        }

        //DTO פונקציה הממירה מרשימה ל
        public static List<OwnerOfEventDto> convertoDtoOwnerOfEventList(List<OwnerOfEvent> p)
        {
            List<OwnerOfEventDto> obj = new List<OwnerOfEventDto>();
            foreach (OwnerOfEvent item in p)
            {
                obj.Add(convertoDtoOwnerOfEvent(item));
            }
            return obj;
        }
        public static List<OwnerOfEvent> convertToTblOwnerOfEventList(List<OwnerOfEventDto> p)
        {
            List<OwnerOfEvent> productTbls = new List<OwnerOfEvent>();
            foreach (OwnerOfEventDto item in p)
            {
                productTbls.Add(convertToTblOwnerOfEvent(item));
            }
            return productTbls;
        }
        //PutInvitedOnTabel

        // DTOפונקציה שממירה מ טמבל☺ ל 
        public static PutInvitedOnTabelDto convertoDtoPutInvitedOnTabel(PutInvitedOnTabel o)
        {
            PutInvitedOnTabelDto obj = new PutInvitedOnTabelDto();
            obj.IdInvitedToEventDto = o.IdInvitedToEvent;
            obj.IdTabelsDto = o.IdTabels;
            obj.IdPutInvitedOnTabelsDto = o.IdPutInvitedOnTabels;
            
            return obj;
        }

        // ☺לטמבל DTOפונקציה שממירה מ  
        public static PutInvitedOnTabel convertToTblPutInvitedOnTabel(PutInvitedOnTabelDto i)
        {
            PutInvitedOnTabel obj = new PutInvitedOnTabel();
            obj.IdInvitedToEvent = i.IdInvitedToEventDto;
            obj.IdTabels = i.IdTabelsDto;
            obj.IdPutInvitedOnTabels = i.IdPutInvitedOnTabelsDto;
           
            return obj;
        }

        //DTO פונקציה הממירה מרשימה ל
        public static List<PutInvitedOnTabelDto> convertoDtoPutInvitedOnTabelList(List<PutInvitedOnTabel> p)
        {
            List<PutInvitedOnTabelDto> obj = new List<PutInvitedOnTabelDto>();
            foreach (PutInvitedOnTabel item in p)
            {
                obj.Add(convertoDtoPutInvitedOnTabel(item));
            }
            return obj;
        }
        public static List<PutInvitedOnTabel> convertToTblPutInvitedOnTabel(List<PutInvitedOnTabelDto> p)
        {
            List<PutInvitedOnTabel> PutInvitedOnTabelTbls = new List<PutInvitedOnTabel>();
            foreach (PutInvitedOnTabelDto item in p)
            {
                PutInvitedOnTabelTbls.Add(convertToTblPutInvitedOnTabel(item));
            }
            return PutInvitedOnTabelTbls;
        }
        //TypeEvent

        // DTOפונקציה שממירה מ טמבל☺ ל 
        public static typeEventDto convertoDtotypeEventTabel(TypeEvent t)
        {
            typeEventDto obj = new typeEventDto();
            obj.IdTypeEventDto = t.IdTypeEvent;
            obj.NameTypeEventDto = t.NameTypeEvent;
            return obj;
        }

        // ☺לטמבל DTOפונקציה שממירה מ  
        public static TypeEvent convertToTbltypeEventTabel(typeEventDto t)
        {
            TypeEvent obj = new TypeEvent();
            obj.IdTypeEvent = t.IdTypeEventDto;
            obj.NameTypeEvent = t.NameTypeEventDto;
            return obj;
        }

        //DTO פונקציה הממירה מרשימה ל
        public static List<typeEventDto> convertoDtoTypeEventList(List<TypeEvent> p)
        {
            List<typeEventDto> obj = new List<typeEventDto>();
            foreach (TypeEvent item in p)
            {
                obj.Add(convertoDtotypeEventTabel(item));
            }
            return obj;
        }
        //tbl פונקציה הממירה מרשימה ל
        public static List<TypeEvent> convertToTblTypeEventlist(List<typeEventDto> p)
        {
            List<TypeEvent> TypeEvents = new List<TypeEvent>();
            foreach (typeEventDto item in p)
            {
                TypeEvents.Add(convertToTbltypeEventTabel(item));
            }
            return TypeEvents;
        }
        //tabalesForEvent

        // DTOפונקציה שממירה מ טמבל☺ ל 
        public static TabelsForEventDto convertoDtotabalesForEventTabel(TabelsForEvent t)
        {
            TabelsForEventDto obj = new TabelsForEventDto();
            obj.IdEventDto = t.IdEvent;
            obj.IdEventNavigationDto = t.IdEventNavigation;
            obj.IdTabelsDto = t.IdTabels;
            obj.IdTypeTabelsDto = t.IdTypeTabels;
            obj.MenOrWomenDto = t.MenOrWomen;
            obj.IdTypeTabelsNavigationDto = t.IdTypeTabelsNavigation;
            return obj;
        }

        // ☺לטמבל DTOפונקציה שממירה מ  
        public static TabelsForEvent convertToTblTabelsForEventTabel(TabelsForEventDto t)
        {
            TabelsForEvent obj = new TabelsForEvent();
            obj.IdTabels = t.IdTabelsDto;
            obj.IdTypeTabels = t.IdTypeTabelsDto;
            obj.IdTypeTabelsNavigation = t.IdTypeTabelsNavigationDto;
            obj.IdEventNavigation = t.IdEventNavigationDto;
            obj.MenOrWomen = t.MenOrWomenDto;
            obj.IdEvent = t.IdEventDto;
            return obj;
        }

        //DTO פונקציה הממירה מרשימה ל
        public static List<TabelsForEventDto> convertoDtoTabelsForEventList(List<TabelsForEvent> p)
        {
            List<TabelsForEventDto> obj = new List<TabelsForEventDto>();
            foreach (TabelsForEvent item in p)
            {
                obj.Add(convertoDtotabalesForEventTabel(item));
            }
            return obj;
        }
        //tbl פונקציה הממירה מרשימה ל
        public static List<TabelsForEvent> convertToTblTabelsForEventlist(List<TabelsForEventDto> p)
        {
            List<TabelsForEvent> TabelsForEvents = new List<TabelsForEvent>();
            foreach (TabelsForEventDto item in p)
            {
                TabelsForEvents.Add(convertToTblTabelsForEventTabel(item));
            }
            return TabelsForEvents;
        }

        //TypeInvite
        public static TypeInviteDto convertoDtoTypeInviteTabel(TypeInvite t)
        {
            TypeInviteDto obj = new TypeInviteDto();
            obj.IdTypeInviteDto = t.IdTypeInvite;
            obj.NameTypeInviteDto = t.NameTypeInvite;
            return obj;
        }

        // ☺לטמבל DTOפונקציה שממירה מ  
        public static TypeInvite convertToTblTypeInviteTabel(TypeInviteDto t)
        {
            TypeInvite obj = new TypeInvite();
            obj.IdTypeInvite = t.IdTypeInviteDto;
            obj.NameTypeInvite = t.NameTypeInviteDto;
            return obj;

        }

        //DTO פונקציה הממירה מרשימה ל
        public static List<TypeInviteDto> convertoDtoTypeInviteList(List<TypeInvite> p)
        {
            List<TypeInviteDto> obj = new List<TypeInviteDto>();
            foreach (TypeInvite item in p)
            {
                obj.Add(convertoDtoTypeInviteTabel(item));
            }
            return obj;
        }
        //tbl פונקציה הממירה מרשימה ל
        public static List<TypeInvite> convertToTblTypeInvitelist(List<TypeInviteDto> p)
        {
            List<TypeInvite> TypeInvite = new List<TypeInvite>();
            foreach (TypeInviteDto item in p)
            {
                TypeInvite.Add(convertToTblTypeInviteTabel(item));
            }
            return TypeInvite;
        }

        //TypeTabels
        public static TypeTabelsDto convertoDtoTypeTabels(TypeTabel t)
        {
            TypeTabelsDto obj = new TypeTabelsDto();
            obj.IdTypeTabelsDto = t.IdTypeTabels;
            obj.NameTypeTabelsDto = t.NameTypeTabels;
            return obj;
        }

        // ☺לטמבל DTOפונקציה שממירה מ  
        public static TypeTabel convertToTblTypeTabelTabel(TypeTabelsDto t)
        {
            TypeTabel obj = new TypeTabel();
            obj.IdTypeTabels = t.IdTypeTabelsDto;
            obj.NameTypeTabels = t.NameTypeTabelsDto;
            return obj;

        }

        //DTO פונקציה הממירה מרשימה ל
        public static List<TypeTabelsDto> convertoDtoTypeTabelsList(List<TypeTabel> p)
        {
            List<TypeTabelsDto> obj = new List<TypeTabelsDto>();
            foreach (TypeTabel item in p)
            {
                obj.Add(convertoDtoTypeTabels(item));
            }
            return obj;
        }
        //tbl פונקציה הממירה מרשימה ל
        public static List<TypeTabel> convertToTblTypeTabellist(List<TypeTabelsDto> p)
        {
            List<TypeTabel> TypeTabels = new List<TypeTabel>();
            foreach (TypeTabelsDto item in p)
            {
                TypeTabels.Add(convertToTblTypeTabelTabel(item));
            }
            return TypeTabels;
        }

    }
}
