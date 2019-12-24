using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;

namespace file_transfer.DTO
{
    public class Teacher
    {
        [Key]
        public int id { get; set; }
        public string name { get; set; }
        public string birthday { get; set; }
        public string university { get; set; }

    }
}
