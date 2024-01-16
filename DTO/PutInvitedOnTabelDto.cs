using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class PutInvitedOnTabelDto
    {
        public int IdPutInvitedOnTabelsDto { get; set; }

        public int? IdTabelsDto { get; set; }

        public int? IdInvitedToEventDto { get; set; }
    }
}
