using System;
using System.Collections.Generic;

namespace DAL.Models;

public partial class TypeEvent
{
    public int IdTypeEvent { get; set; }

    public string? NameTypeEvent { get; set; }

    public virtual ICollection<OwnerOfEvent> OwnerOfEvents { get; } = new List<OwnerOfEvent>();
}
