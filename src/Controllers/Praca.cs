using System.Collections.Generic;

namespace ExcelExportExample.Controllers
{
    public class Praca
    {
        public Praca()
        {
            Passagens = new List<Passagem>();
        }

        public string Nome { get; set; }
        public IEnumerable<Passagem> Passagens { get; set; }
    }
}
