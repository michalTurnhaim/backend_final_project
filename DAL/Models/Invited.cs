using System;
using System.Collections.Generic;

namespace DAL.Models;

public partial class Invited
{
    public string EmailInvited { get; set; } = null!;

    public string? FirstNameInvited { get; set; }

    public string? LastNameInvited { get; set; }

    public int? PassWord { get; set; }

    public virtual ICollection<InvitedToEvent> InvitedToEvents { get; } = new List<InvitedToEvent>();

    public virtual ICollection<OwnerOfEvent> OwnerOfEvents { get; } = new List<OwnerOfEvent>();
}
