using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CNC_Run_Times_and_Material_Counts
{
    class Material
    {
        public Material (string name, int quantity)
        {
            this.Name = name;
            this.Quantity = quantity;
        }
        public Material(string name)
        {
            this.Name = name;
            this.Quantity = 0;
        }

        public string Name { get; set; }
        public int Quantity { get; set; }
    }
}
