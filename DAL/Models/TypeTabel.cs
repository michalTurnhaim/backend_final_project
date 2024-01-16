using System;
using System.Collections.Generic;

namespace DAL.Models;

public partial class TypeTabel
{
    public int IdTypeTabels { get; set; }

    public string? NameTypeTabels { get; set; }

    public virtual ICollection<TabelsForEvent> TabelsForEvents { get; } = new List<TabelsForEvent>();
}
