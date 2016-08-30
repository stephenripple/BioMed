using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BioMed.Data.Models
{
    public class BioMedProduct
    {
        public string Make;
        public string Model;
        public string Features;
        public List<BioMedProductSpec> Specs;
    }
}
