using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace NfeXmlToExcel
{
    public static class NfeParser
    {
        private static readonly XNamespace Ns = "http://www.portalfiscal.inf.br/nfe";

        public static NfeDocument ParseFromFile(string filePath)
        {
            var xdoc = XDocument.Load(filePath, LoadOptions.PreserveWhitespace | LoadOptions.SetLineInfo);

            // Pode vir como <nfeProc>...<NFe>...</NFe>...</nfeProc> ou direto <NFe>
            var nfe = xdoc.Descendants(Ns + "NFe").FirstOrDefault()
          ?? ((xdoc.Root != null && xdoc.Root.Name.LocalName == "NFe") ? xdoc.Root : null);

if (nfe == null)
    throw new Exception("Não encontrei a tag NFe no XML.");

            var infNFe = nfe.Descendants(Ns + "infNFe").FirstOrDefault();
            if (infNFe == null)
                throw new Exception("Não encontrei a tag infNFe no XML.");

            var chaveId = (string?)infNFe.Attribute("Id") ?? "";
            var chave = chaveId.StartsWith("NFe", StringComparison.OrdinalIgnoreCase) ? chaveId.Substring(3) : chaveId;

            var ide = infNFe.Element(Ns + "ide");
            var emit = infNFe.Element(Ns + "emit");
            var enderEmit = emit?.Element(Ns + "enderEmit");

            var dest = infNFe.Element(Ns + "dest");
            var enderDest = dest?.Element(Ns + "enderDest");

            var total = infNFe.Element(Ns + "total");
            var icmsTot = total?.Element(Ns + "ICMSTot");

            // Protocolo (quando XML for nfeProc)
            var prot = xdoc.Descendants(Ns + "protNFe").FirstOrDefault();
            var infProt = prot?.Descendants(Ns + "infProt").FirstOrDefault();

            var header = new NfeHeader
            {
                FileName = Path.GetFileName(filePath),
                Chave = chave,

                Numero = ide?.Element(Ns + "nNF")?.Value ?? "",
                Serie = ide?.Element(Ns + "serie")?.Value ?? "",
                DhEmi = ParseDate(ide?.Element(Ns + "dhEmi")?.Value ?? ide?.Element(Ns + "dEmi")?.Value),

                NatOp = ide?.Element(Ns + "natOp")?.Value ?? "",
                TipoNF = ide?.Element(Ns + "tpNF")?.Value ?? "",

                EmitCNPJ = emit?.Element(Ns + "CNPJ")?.Value ?? "",
                EmitNome = emit?.Element(Ns + "xNome")?.Value ?? "",
                EmitUF = enderEmit?.Element(Ns + "UF")?.Value ?? "",
                EmitMun = enderEmit?.Element(Ns + "xMun")?.Value ?? "",

                DestDoc = dest?.Element(Ns + "CNPJ")?.Value
                          ?? dest?.Element(Ns + "CPF")?.Value
                          ?? "",
                DestNome = dest?.Element(Ns + "xNome")?.Value ?? "",
                DestUF = enderDest?.Element(Ns + "UF")?.Value ?? "",

                VProd = ParseDec(icmsTot?.Element(Ns + "vProd")?.Value),
                VNF = ParseDec(icmsTot?.Element(Ns + "vNF")?.Value),
                VICMS = ParseDec(icmsTot?.Element(Ns + "vICMS")?.Value),
                VIPI = ParseDec(icmsTot?.Element(Ns + "vIPI")?.Value),
                VPIS = ParseDec(icmsTot?.Element(Ns + "vPIS")?.Value),
                VCOFINS = ParseDec(icmsTot?.Element(Ns + "vCOFINS")?.Value),
                VFrete = ParseDec(icmsTot?.Element(Ns + "vFrete")?.Value),
                VDesc = ParseDec(icmsTot?.Element(Ns + "vDesc")?.Value),
                VOutro = ParseDec(icmsTot?.Element(Ns + "vOutro")?.Value),
                VTotTrib = ParseDec(icmsTot?.Element(Ns + "vTotTrib")?.Value),

                ProtNFe = infProt?.Element(Ns + "nProt")?.Value ?? "",
                DhRecbto = ParseDate(infProt?.Element(Ns + "dhRecbto")?.Value),
                StatusProt = infProt?.Element(Ns + "cStat")?.Value ?? ""
            };

            var doc = new NfeDocument { Header = header };

            var dets = infNFe.Elements(Ns + "det").ToList();
            foreach (var det in dets)
            {
                int? nItem = ParseInt(det.Attribute("nItem")?.Value);

                var prod = det.Element(Ns + "prod");
                var imposto = det.Element(Ns + "imposto");

                // ICMS é um nó que tem "um tipo dentro" (ICMS00, ICMSSN102, etc.)
                var icmsNode = imposto?.Element(Ns + "ICMS")?.Elements().FirstOrDefault();
                var cst = icmsNode?.Element(Ns + "CST")?.Value
                          ?? icmsNode?.Element(Ns + "CSOSN")?.Value
                          ?? "";

                var item = new NfeItem
                {
                    Chave = chave,
                    NItem = nItem,

                    CProd = prod?.Element(Ns + "cProd")?.Value ?? "",
                    XProd = prod?.Element(Ns + "xProd")?.Value ?? "",
                    NCM = prod?.Element(Ns + "NCM")?.Value ?? "",
                    CFOP = prod?.Element(Ns + "CFOP")?.Value ?? "",

                    UCom = prod?.Element(Ns + "uCom")?.Value ?? "",
                    QCom = ParseDec(prod?.Element(Ns + "qCom")?.Value),
                    VUnCom = ParseDec(prod?.Element(Ns + "vUnCom")?.Value),
                    VProd = ParseDec(prod?.Element(Ns + "vProd")?.Value),

                    CST_CSOSN = cst,
                    PICMS = ParseDec(icmsNode?.Element(Ns + "pICMS")?.Value),
                    VICMS = ParseDec(icmsNode?.Element(Ns + "vICMS")?.Value),

                    VIPI = ParseDec(imposto?.Descendants(Ns + "IPI").Descendants(Ns + "vIPI").FirstOrDefault()?.Value),
                    VPIS = ParseDec(imposto?.Descendants(Ns + "PIS").Descendants(Ns + "vPIS").FirstOrDefault()?.Value),
                    VCOFINS = ParseDec(imposto?.Descendants(Ns + "COFINS").Descendants(Ns + "vCOFINS").FirstOrDefault()?.Value),
                };

                doc.Items.Add(item);
            }

            return doc;
        }

        private static decimal? ParseDec(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (decimal.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var v))
                return v;
            return null;
        }

        private static int? ParseInt(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (int.TryParse(s.Trim(), out var v)) return v;
            return null;
        }

        private static DateTime? ParseDate(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return null;

            // dhEmi costuma vir ISO 8601. dEmi pode vir yyyy-MM-dd
            if (DateTimeOffset.TryParse(s.Trim(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dto))
                return dto.DateTime;

            if (DateTime.TryParse(s.Trim(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt))
                return dt;

            return null;
        }
    }
}
