using System;
using System.Collections.Generic;

namespace DAL.Models;

public partial class PutInvitedOnTabel
{
    public int IdPutInvitedOnTabels { get; set; }

    public int? IdTabels { get; set; }

    public int? IdInvitedToEvent { get; set; }

    public virtual InvitedToEvent? IdInvitedToEventNavigation { get; set; }

    public virtual TabelsForEvent? IdTabelsNavigation { get; set; }
}
