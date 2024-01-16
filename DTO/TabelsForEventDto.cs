using DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class TabelsForEventDto
    {
        public int IdTabelsDto { get; set; }

        public bool? MenOrWomenDto { get; set; }

        public int? IdEventDto { get; set; }

        public int? IdTypeTabelsDto { get; set; }

        public virtual OwnerOfEvent? IdEventNavigationDto { get; set; }

        public virtual TypeTabel? IdTypeTabelsNavigationDto { get; set; }
    }
}
