using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL.Models;
using DAL.action;
using DTO;

namespace BLL
{
    public class InvitedBLL : IInvitedBLL
    {
        IInvitedAction _invitedAction;
        public InvitedBLL(IInvitedAction invitedAction)
        {
            _invitedAction = invitedAction;
        }
        // מקבל את כל המוזמנים
        public List<InvitedDto> getAllInvited()
        {
            List<Invited> list = _invitedAction.GetAllInvited();
            List<InvitedDto> Inviteds = new List<InvitedDto>();

            foreach (Invited item in list)
            {
                InvitedDto i = Converts.convertoDtoInvited(item);
                Inviteds.Add(i);
            }
            return Inviteds;
        }

        ////מחזיר מוצר לפי קוד
        //public static InvitedDto getProductById(int id)
        //{
        //    ProductsTbl p = ProductsDAL.getProductsFromTblById(id);
        //    ProductDTO newProd = IMapper.Map<ProductsTbl, ProductDTO>(p);
        //    return newProd;
        //}

        ////מחזיר מוצרים לפי קוד קטגוריה
        //public static List<ProductDTO> getProductByCategoryId(int id)
        //{
        //    List<ProductsTbl> list = ProductsDAL.getProductByCategoryId(id);
        //    List<ProductDTO> products = new List<ProductDTO>();

        //    foreach (ProductsTbl item in list)
        //    {
        //        products.Add(IMapper.Map<ProductsTbl, ProductDTO>(item));
        //    }
        //    return products;
        //}

        //הוספת מוזמן לארוע
        public List<InvitedDto> addInvited(InvitedDto i)
        {
            //if (p.CategoryName == "")
            //    p.CategoryName = CategoryDAL.getCategoriesFromTblById(p.CategoryId).CategoryName;

            Invited Inviteds = Converts.convertTotblInvuted(i);
            List<Invited> list = _invitedAction.AddInvited(Inviteds);
            return getAllInvited();
        }

        //מחיקת מוזמן לארוע
        public List<InvitedDto> deleteInvited(string id)
        {
            List<Invited> list = _invitedAction.dellInvited(id);
            return getAllInvited();
        }
        //עדכון מוזמן לארוע
        public List<InvitedDto> updateInvited(InvitedDto i, string id)
        {
            Invited invitedToUpdate = Converts.convertTotblInvuted(i);
            List<Invited> newList = _invitedAction.updateInvited(id, invitedToUpdate);
            return getAllInvited();
        }
        //התחברותת
        public InvitedDto login(string email, int password)
        {
            Invited j = _invitedAction.GetAllInvited().FirstOrDefault(n => n.EmailInvited.Equals(email) && n.PassWord == password);
            if (j == null)
                return null;
            InvitedDto p = Converts.convertoDtoInvited(j);

            return p;
        }
        //בדיקה אם במייל קיים במערכת
        public InvitedDto checEmailIfExists(string email)
        {
            Invited j = _invitedAction.GetAllInvited().FirstOrDefault(n => n.EmailInvited.Equals(email) );
            if (j == null)
                return null;
            InvitedDto p = Converts.convertoDtoInvited(j);

            return p;
        }

    }
}
