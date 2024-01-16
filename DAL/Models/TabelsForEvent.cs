using System;
using System.Collections.Generic;

namespace DAL.Models;

public partial class TabelsForEvent
{
    public int IdTabels { get; set; }

    public bool? MenOrWomen { get; set; }

    public int? IdEvent { get; set; }

    public int? IdTypeTabels { get; set; }

    public virtual OwnerOfEvent? IdEventNavigation { get; set; }

    public virtual TypeTabel? IdTypeTabelsNavigation { get; set; }

    public virtual ICollection<PutInvitedOnTabel> PutInvitedOnTabels { get; } = new List<PutInvitedOnTabel>();
}
