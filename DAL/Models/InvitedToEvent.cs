using System;
using System.Collections.Generic;

namespace DAL.Models;

public partial class InvitedToEvent
{
    public int IdInvitedToEvent { get; set; }

    public string? EmailInvited { get; set; }

    public int? IdEvent { get; set; }

    public int? IdTypeInvite { get; set; }

    public bool? IsCome { get; set; }

    public int? NumBoys { get; set; }

    public int? NumSonAdults { get; set; }

    public int? NumGirls { get; set; }

    public int? NumTeenageGirls { get; set; }

    public int? NumDaughterAdults { get; set; }

    public int? NumteenageBoys { get; set; }

    public virtual Invited? EmailInvitedNavigation { get; set; }

    public virtual OwnerOfEvent? IdEventNavigation { get; set; }

    public virtual TypeInvite? IdTypeInviteNavigation { get; set; }

    public virtual ICollection<PutInvitedOnTabel> PutInvitedOnTabels { get; } = new List<PutInvitedOnTabel>();
}
