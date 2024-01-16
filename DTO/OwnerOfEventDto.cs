using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class OwnerOfEventDto
    {
        public int IdEventDto { get; set; }

        public string? FirstNameOwnerOfEventDto { get; set; }

        public string? LastNameOwnerOfEventDto { get; set; }

        public string? EmailOwnerOfEventDto { get; set; }

        public DateTime? DateOfEventDto { get; set; }

        public string? AdressOfEventDto { get; set; }

        public int? IdTypeEventDto { get; set; }

        public string? NameFileInvitationDto { get; set; }
    }
}
