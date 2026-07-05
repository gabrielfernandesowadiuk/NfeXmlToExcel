using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NfeXmlToExcel
{
    public static class ExcelExporter
    {
        public static void Export(string outputPath, List<NfeHeader> headers, List<NfeItem> items, List<ParseError> errors)
        {
            using var wb = new XLWorkbook();

            // Aba principal no modelo do seu arquivo
            WriteCenarioSheetLikeModel(wb, headers, items);

            // (Opcional) abas auxiliares que ajudam muito a conferir
            WriteHeadersSheet(wb, headers);
            WriteItemsSheet(wb, items);
            WriteErrorsSheet(wb, errors);

            wb.SaveAs(outputPath);
        }

        // =========================
        // ✅ ABA CENÁRIO (MODELO)
        //  - 65 colunas
        //  - header na linha 4
        //  - dados na linha 5+
        //  - 1 linha por ITEM (det)
        // =========================
        private static void WriteCenarioSheetLikeModel(XLWorkbook wb, List<NfeHeader> headers, List<NfeItem> items)
        {
            if (wb.Worksheets.Any(s => s.Name == "Cenário"))
                wb.Worksheets.Delete("Cenário");

            var ws = wb.AddWorksheet("Cenário");

            // Igual ao modelo: texto na linha 3 coluna A
            ws.Cell(3, 1).SetValue("Soma vase");

            var headerRow = 4;
            var dataRowStart = 5;

            // 65 colunas iguais ao seu modelo
            string[] cols =
            {
                "Dt Cancela",
                "Fat",
                "Nat Oper",
                "Denominação",
                "Emissão",
                "Ser",
                "Nota Fis",
                "Cliente/Fornec",
                "NomeCliente/Fornec",
                "Ins Estadual",
                "Item",
                "Descrição Item",
                "Class Fiscal",
                "Peso Brut",
                "Pre Liq",
                "Vl Mercad Liq",
                "Peso Liq",
                "Vl Tot Item",
                "Retencao COFINS",
                "Retencao PIS",
                "IRRF",
                "Retencao CSLL",
                "% ICMS",
                "Base ICMS Item",
                "Vl ICMS Item",
                "Vl ICMS Não Trib",
                "Vl ICMS Outras",
                "Vl ICMS Subst",
                "Trib ICM",
                "Base IPI Item",
                "%RICMS",
                "Vl IPI",
                "Vl IPI Outras",
                "Percentual Icms Dife",
                "Vl IPI Não Trib",
                "Sit Trib IPI",
                "% RIPI",
                "Trib IPI",
                "% IPI",
                "Sit Trib COFINS",
                "Sit Trib PIS",
                "Vl ISS",
                "Base ISS Item",
                "Vl ISS Não Trib",
                "Vl ISS Outras",
                "% ISS",
                "Trib ISS",
                "% RISS",
                "GE",
                "Descrição GE",
                "Família",
                "Descrição Fam",
                "Fam Coml",
                "Desc",
                "Item fiscal",
                "Base COFINS fisc",
                "COFINS fisc",
                "% Cofins",
                "Base PIS fisc",
                "Vl PIS fisc",
                "% Pis",
                "Qt",
                "Hr Nota Fis",
                "UF",
                "Total"
            };

            // Header
            for (int c = 0; c < cols.Length; c++)
                ws.Cell(headerRow, c + 1).SetValue(cols[c]);

            ws.SheetView.FreezeRows(headerRow);

            // Mapa para achar o cabeçalho pela chave
            var headerByChave = headers
                .GroupBy(h => h.Chave)
                .ToDictionary(g => g.Key, g => g.First());

            // ✅ 1 linha por item
            int r = dataRowStart;

            foreach (var it in items.OrderBy(i => i.Chave).ThenBy(i => i.NItem ?? 0))
            {
                headerByChave.TryGetValue(it.Chave, out var h);

                // ✅ OBRIGATÓRIO (o que você pediu):
                SetCell(ws.Cell(r, 11), it.CProd); // Item (cProd)
                SetCell(ws.Cell(r, 12), it.XProd); // Descrição Item (xProd)

                // ✅ Campos que ajudam muito (baseado no XML):
                SetCell(ws.Cell(r, 13), it.NCM);   // Class Fiscal
                SetCell(ws.Cell(r, 15), it.VUnCom); // Pre Liq
                SetCell(ws.Cell(r, 18), it.VProd); // Vl Tot Item
                SetCell(ws.Cell(r, 62), it.QCom);  // Qt

                // Dados básicos da NF (úteis e normalmente existem no XML):
                SetCell(ws.Cell(r, 3), h?.NatOp);   // Nat Oper
                SetCell(ws.Cell(r, 5), h?.DhEmi);   // Emissão
                SetCell(ws.Cell(r, 6), h?.Serie);   // Ser
                SetCell(ws.Cell(r, 7), h?.Numero);  // Nota Fis

                // Em NF-e de ENTRADA: "Cliente/Fornec" normalmente é o FORNECEDOR (emitente)
                SetCell(ws.Cell(r, 8), h?.EmitCNPJ);
                SetCell(ws.Cell(r, 9), h?.EmitNome);
                SetCell(ws.Cell(r, 64), h?.EmitUF);

                // ICMS / IPI / PIS / COFINS por item (quando existir no XML)
                SetCell(ws.Cell(r, 23), it.PICMS);  // % ICMS
                SetCell(ws.Cell(r, 25), it.VICMS);  // Vl ICMS Item
                SetCell(ws.Cell(r, 24), EstimateBaseFromValueAndRate(it.VICMS, it.PICMS)); // Base ICMS (estimada)

                SetCell(ws.Cell(r, 32), it.VIPI);   // Vl IPI
                SetCell(ws.Cell(r, 60), it.VPIS);   // Vl PIS fisc
                SetCell(ws.Cell(r, 57), it.VCOFINS);// COFINS fisc

                // Total da NF (repete em todas as linhas do item)
                SetCell(ws.Cell(r, 65), h?.VNF);    // Total

                r++;
            }

            // Como tabela + autofilter (ficar igual planilha de trabalho)
            var used = ws.RangeUsed();
            if (used != null)
            {
                used.CreateTable();
                ws.Columns().AdjustToContents();
            }
        }

        private static decimal? EstimateBaseFromValueAndRate(decimal? taxValue, decimal? ratePercent)
        {
            if (taxValue is null || ratePercent is null) return null;
            if (ratePercent <= 0) return null;
            return Math.Round(taxValue.Value / (ratePercent.Value / 100m), 2);
        }

        // =========================
        // Abas auxiliares (opcionais)
        // =========================
        private static void WriteHeadersSheet(XLWorkbook wb, List<NfeHeader> headers)
        {
            var ws = wb.AddWorksheet("NFe");

            var cols = new (string Title, Func<NfeHeader, object?> Value)[]
            {
                ("Arquivo", x => x.FileName),
                ("Chave", x => x.Chave),
                ("Número", x => x.Numero),
                ("Série", x => x.Serie),
                ("Emissão", x => x.DhEmi),
                ("Natureza Operação", x => x.NatOp),
                ("Tipo NF", x => x.TipoNF),

                ("Emit CNPJ", x => x.EmitCNPJ),
                ("Emit Nome", x => x.EmitNome),
                ("Emit UF", x => x.EmitUF),
                ("Emit Município", x => x.EmitMun),

                ("Dest Doc", x => x.DestDoc),
                ("Dest Nome", x => x.DestNome),
                ("Dest UF", x => x.DestUF),

                ("vProd", x => x.VProd),
                ("vNF", x => x.VNF),
                ("vICMS", x => x.VICMS),
                ("vIPI", x => x.VIPI),
                ("vPIS", x => x.VPIS),
                ("vCOFINS", x => x.VCOFINS),
                ("vFrete", x => x.VFrete),
                ("vDesc", x => x.VDesc),
                ("vOutro", x => x.VOutro),
                ("vTotTrib", x => x.VTotTrib),
            };

            for (int c = 0; c < cols.Length; c++)
                ws.Cell(1, c + 1).SetValue(cols[c].Title);

            for (int r = 0; r < headers.Count; r++)
                for (int c = 0; c < cols.Length; c++)
                    SetCell(ws.Cell(r + 2, c + 1), cols[c].Value(headers[r]));

            ws.SheetView.FreezeRows(1);

            var used = ws.RangeUsed();
            if (used != null)
            {
                used.CreateTable();
                ws.Columns().AdjustToContents();
            }
        }

        private static void WriteItemsSheet(XLWorkbook wb, List<NfeItem> items)
        {
            var ws = wb.AddWorksheet("Itens");

            var cols = new (string Title, Func<NfeItem, object?> Value)[]
            {
                ("Chave", x => x.Chave),
                ("nItem", x => x.NItem),
                ("cProd", x => x.CProd),
                ("xProd", x => x.XProd),
                ("NCM", x => x.NCM),
                ("CFOP", x => x.CFOP),
                ("uCom", x => x.UCom),
                ("qCom", x => x.QCom),
                ("vUnCom", x => x.VUnCom),
                ("vProd", x => x.VProd),
                ("CST/CSOSN", x => x.CST_CSOSN),
                ("pICMS", x => x.PICMS),
                ("vICMS", x => x.VICMS),
                ("vIPI", x => x.VIPI),
                ("vPIS", x => x.VPIS),
                ("vCOFINS", x => x.VCOFINS),
            };

            for (int c = 0; c < cols.Length; c++)
                ws.Cell(1, c + 1).SetValue(cols[c].Title);

            for (int r = 0; r < items.Count; r++)
                for (int c = 0; c < cols.Length; c++)
                    SetCell(ws.Cell(r + 2, c + 1), cols[c].Value(items[r]));

            ws.SheetView.FreezeRows(1);

            var used = ws.RangeUsed();
            if (used != null)
            {
                used.CreateTable();
                ws.Columns().AdjustToContents();
            }
        }

        //testes
        private static void WriteErrorsSheet(XLWorkbook wb, List<ParseError> errors)
        {
            var ws = wb.AddWorksheet("Erros");
            ws.Cell(1, 1).SetValue("Arquivo");
            ws.Cell(1, 2).SetValue("Erro");

            for (int r = 0; r < errors.Count; r++)
            {
                ws.Cell(r + 2, 1).SetValue(errors[r].FilePath);
                ws.Cell(r + 2, 2).SetValue(errors[r].Message);
            }

            ws.SheetView.FreezeRows(1);

            var used = ws.RangeUsed();
            if (used != null)
            {
                used.CreateTable();
                ws.Columns().AdjustToContents();
            }
        }

        private static void SetCell(IXLCell cell, object? value)
        {
            if (value == null)
            {
                cell.Clear();
                return;
            }

            switch (value)
            {
                case string s: cell.SetValue(s); break;
                case int i: cell.SetValue(i); break;
                case long l: cell.SetValue(l); break;
                case double d: cell.SetValue(d); break;
                case decimal dec: cell.SetValue(dec); break;

                case DateTime dt:
                    cell.SetValue(dt);
                    cell.Style.DateFormat.Format = "yyyy-mm-dd";
                    break;

                case DateTimeOffset dto:
                    cell.SetValue(dto.DateTime);
                    cell.Style.DateFormat.Format = "yyyy-mm-dd";
                    break;

                case bool b: cell.SetValue(b); break;

                default:
                    cell.SetValue(value.ToString() ?? "");
                    break;
            }
        }
    }
}
