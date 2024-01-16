using System;
using System.Collections.Generic;

namespace DAL.Models;

public partial class OwnerOfEvent
{
    public int IdEvent { get; set; }

    public string? EmailOwnerOfEvent { get; set; }

    public DateTime? DateOfEvent { get; set; }

    public string? AdressOfEvent { get; set; }

    public int? IdTypeEvent { get; set; }

    public string? NameFileInvitation { get; set; }

    public virtual Invited? EmailOwnerOfEventNavigation { get; set; }

    public virtual TypeEvent? IdTypeEventNavigation { get; set; }

    public virtual ICollection<InvitedToEvent> InvitedToEvents { get; } = new List<InvitedToEvent>();

    public virtual ICollection<TabelsForEvent> TabelsForEvents { get; } = new List<TabelsForEvent>();
}
