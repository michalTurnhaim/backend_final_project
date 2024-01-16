using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class InvitedToEventDto
    {
        public int IdInvitedToEventDto { get; set; }

        public string? EmailInvitedDto { get; set; }

        public int? IdEventDto { get; set; }
        //משתנים שהוספנו
        public string? pic { get; set; }
        public string? firstNameOwner { get; set; }
        public string? lastNameOwner { get; set; }

        public DateTime? DateOfEventDto { get; set; }

        public string? FName { get; set; }

        public string? LName { get; set; }

        public int? IdTypeInviteDto { get; set; }

        public int? NumSonAdultsDto { get; set; }

        public int? NumDaughterAdultsDto { get; set; }

        public int? NumteenageBoysDto { get; set; }

        public int? NumTeenageGirlsDto { get; set; }

        public int? NumBoysDto { get; set; }

        public int? NumGirlsDto { get; set; }
        public bool? IsComeDto { get; set; }
    }
}
