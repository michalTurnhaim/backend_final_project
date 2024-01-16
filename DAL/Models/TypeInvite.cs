using System;
using System.Collections.Generic;

namespace DAL.Models;

public partial class TypeInvite
{
    public int IdTypeInvite { get; set; }

    public string? NameTypeInvite { get; set; }

    public virtual ICollection<InvitedToEvent> InvitedToEvents { get; } = new List<InvitedToEvent>();
}
