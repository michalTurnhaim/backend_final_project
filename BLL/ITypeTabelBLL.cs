using DTO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BLL
{
    public interface ITypeTabelBLL
    {

        public List<TypeTabelsDto> getAllTypeTabel();
        public List<TypeTabelsDto> addTypeTabel(TypeTabelsDto i);
        public List<TypeTabelsDto> deleteTypeTabel(int id);
        public List<TypeTabelsDto> updateTypeTabel(TypeTabelsDto i, int id);
    }
}
