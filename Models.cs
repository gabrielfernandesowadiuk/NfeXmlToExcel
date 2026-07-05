using System;
using System.Collections.Generic;

namespace NfeXmlToExcel
{
    public class NfeDocument
    {
        public NfeHeader Header { get; set; } = new();
        public List<NfeItem> Items { get; set; } = new();
    }

    public class NfeHeader
    {
        public string FileName { get; set; } = "";
        public string Chave { get; set; } = "";

        public string Numero { get; set; } = "";
        public string Serie { get; set; } = "";
        public DateTime? DhEmi { get; set; }

        public string NatOp { get; set; } = "";
        public string TipoNF { get; set; } = ""; // 0 entrada / 1 saída

        public string EmitCNPJ { get; set; } = "";
        public string EmitNome { get; set; } = "";
        public string EmitUF { get; set; } = "";
        public string EmitMun { get; set; } = "";

        public string DestDoc { get; set; } = ""; // CNPJ/CPF
        public string DestNome { get; set; } = "";
        public string DestUF { get; set; } = "";

        public decimal? VProd { get; set; }
        public decimal? VNF { get; set; }
        public decimal? VICMS { get; set; }
        public decimal? VIPI { get; set; }
        public decimal? VPIS { get; set; }
        public decimal? VCOFINS { get; set; }
        public decimal? VFrete { get; set; }
        public decimal? VDesc { get; set; }
        public decimal? VOutro { get; set; }
        public decimal? VTotTrib { get; set; }

        public string ProtNFe { get; set; } = "";
        public DateTime? DhRecbto { get; set; }
        public string StatusProt { get; set; } = "";
    }

    public class NfeItem
    {
        public string Chave { get; set; } = "";
        public int? NItem { get; set; }

        public string CProd { get; set; } = "";
        public string XProd { get; set; } = "";
        public string NCM { get; set; } = "";
        public string CFOP { get; set; } = "";

        public string UCom { get; set; } = "";
        public decimal? QCom { get; set; }
        public decimal? VUnCom { get; set; }
        public decimal? VProd { get; set; }

        public string CST_CSOSN { get; set; } = "";
        public decimal? PICMS { get; set; }
        public decimal? VICMS { get; set; }

        public decimal? VIPI { get; set; }
        public decimal? VPIS { get; set; }
        public decimal? VCOFINS { get; set; }
    }

    public class ParseError
    {
        public string FilePath { get; set; } = "";
        public string Message { get; set; } = "";
    }
}
