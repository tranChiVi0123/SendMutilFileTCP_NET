using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;

namespace file_transfer.DTO
{
    public class Room
    {
        [Key]
        public int id { get; set; }
        public string nameRoom { get; set; }
        public string localtion { get; set; }
        public string note { get; set; }

    }
}
