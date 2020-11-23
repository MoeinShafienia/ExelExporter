using System.Collections.Generic;
using GenFu;

namespace ExelExporter
{
    class ProductFactory
    {
        public IList<Product> GetListOfProdutsByGenfu()
        {
            return A.ListOf<Product>(20);
        }
    }
}