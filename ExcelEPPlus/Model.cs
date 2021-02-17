using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace ExcelEPPlus
{
    public class Model
    {
        [OrderByAttribute(1)]
        [DisplayName("編號")]
        public int Id { get; set; }
    }
}
