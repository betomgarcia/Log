using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Log.Model
{
    public class NationalPurchaseSync
    {
        public string? IdNotaScott { get; set; }
        public string? IdeProdutoScott { get; set; }

        public string? NumeroNota { get; set; }

        public string? Filial { get; set; }
        public string? Tes { get; set; }
        public string? CondicaoPagamento { get; set; }
        public string? Natureza { get; set; }

        public string? Fornecedor { get; set; }

        public string? TipoErro { get; set; }
        public string? Error { get; set; }
    }
}
