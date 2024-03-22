using CsvHelper;
using Log.Model;
using OfficeOpenXml;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Log.Convert
{
    public static class ConvertCsv
    {
        public static void ConvertErrosProductEnricher(Stream file)
        {
            if (file is not null)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int cols = worksheet.Dimension.End.Column;
                    int rows = worksheet.Dimension.End.Row;

                    var listToCsv = new List<Enricher>();

                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= cols; j++)
                        {
                            if (worksheet.Cells[i, j] != null && worksheet.Cells[i, j].Value != null)
                            {
                                var cellValue = worksheet.Cells[i, j].Value.ToString();
                                var linha = "campo nao encontrado na linha A" + i;

                                if (cellValue!.Contains("errorMessage"))
                                {
                                    var splitedValue = cellValue.Split("Message:");

                                    if (splitedValue.Length > 0)
                                    {
                                        var messageJson = splitedValue.Count() > 1 ? splitedValue[1] : splitedValue[0];

                                        var idScott = ConvertFieldJustNumber(messageJson, "B1_XIDEPRD", linha);
                                        var partNumberOriginal = ConvertFieldModelo(messageJson);
                                        var ncm = ConvertFieldJustNumber(messageJson, "B1_POSIPI", linha);

                                        var errosMsgSplit = messageJson.Split("errorMessage\"\":\"\"\\\"\"");
                                        var errosMsgKey = errosMsgSplit.Count() > 1 ? errosMsgSplit[1] : errosMsgSplit[0];
                                        var errosMsgValue = errosMsgKey.Substring(0, errosMsgKey.IndexOf("\\\"\"\"\",\"\""));
                                        var error = errosMsgValue;

                                        listToCsv.Add(new Enricher()
                                        {
                                            idScott = idScott.Trim(),
                                            partNumberOriginal = partNumberOriginal.Trim(),
                                            ncm = ncm,
                                            error = error.ToUTF8()
                                        });

                                    }
                                }
                            }
                        }
                    }

                    var syncs = listToCsv.Distinct().ToList();

                    var dataStg = DateTime.Now.ToString();
                    dataStg = Regex.Replace(dataStg, "[^0-9a-zA-Z]+", "");


                    using (var writer = new StreamWriter(@$"C:\Users\rogarcia\OneDrive - ScanSource, Inc\Documentos\ErrosLog\ErrosProductEnricherConsumer" + dataStg + ".csv"))

                    using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                    {
                        csv.WriteRecords(syncs);
                    }

                    Console.WriteLine("Fim da criação do CSV do Enricher");
                }
            }
        }
        public static void ConvertErrosProductSync(Stream file)
        {
            if (file is not null)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int cols = worksheet.Dimension.End.Column;
                    int rows = worksheet.Dimension.End.Row;

                    var listToCsv = new List<Sync>();

                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= cols; j++)
                        {
                            if (worksheet.Cells[i, j] != null && worksheet.Cells[i, j].Value != null)
                            {
                                var linha = "campo nao encontrado na linha A" + i;
                                var cellValue = worksheet.Cells[i, j].Value.ToString();

                                if (cellValue!.Contains("errorMessage"))
                                {
                                    var splitedValue = cellValue.Split("Message:");

                                    if (splitedValue.Length > 0)
                                    {
                                        var messageJson = splitedValue.Count() > 1 ? splitedValue[1] : splitedValue[0];

                                        var descricaoErro = ConvertFieldError(messageJson);
                                        var tipoErroMsgValue = ConvertFieldTipoErro(descricaoErro);
                                        var idScott = ConvertFieldJustNumber(messageJson, "PartNumberId", linha);
                                        var partNumberOriginal = ConvertFieldWithReplaceOriginalPartNumber(messageJson, "OriginalPartNumber", linha);
                                        var partNumberInterno = ConvertFieldWithReplace(messageJson, "InternalPartNumber", linha);

                                        var ncm = ConvertFieldJustNumber(messageJson, "TaxClassificationCode", linha);
                                        ncm = ncm.Equals("nullPRDCLFIDE0") ? string.Empty : ncm;

                                        var intangiblencm = ConvertFieldJustNumber(messageJson, "IntangibleTaxClassificationCode", linha);
                                        intangiblencm = intangiblencm.Equals("nullPRDCLFIDE0") ? string.Empty : intangiblencm;

                                        var filial = ConvertFieldJustNumber(messageJson, "BillingLocation", linha);

                                        var fabric = ConvertFieldWithReplace(messageJson, "ManufacturerName", linha);

                                        listToCsv.Add(new Sync()
                                        {
                                            IdScott = idScott.Trim(),
                                            PartNumberOriginal = partNumberOriginal.Trim(),
                                            Ncm = ncm,
                                            IntangibleNcm = intangiblencm,
                                            Filial = filial.Trim(),
                                            Error = descricaoErro.ToUTF8(),
                                            Fabrica = fabric.Trim(),
                                            PartNumberInterno = partNumberInterno.Trim(),
                                            TipoErro = tipoErroMsgValue.Trim().ToUTF8()
                                        });

                                    }
                                }
                            }
                        }
                    }

                    var syncs = listToCsv.Distinct().ToList();

                    var dataStg = DateTime.Now.ToString();
                    dataStg = Regex.Replace(dataStg, "[^0-9a-zA-Z]+", "");

                    using (var writer = new StreamWriter(@"C:\Users\rogarcia\OneDrive - ScanSource, Inc\Documentos\ErrosLog\ProtheusProductSyncConsumer" + dataStg + ".csv"))

                    using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
                    {
                        csv.WriteRecords(syncs);
                    }

                    Console.WriteLine("Fim da criação do CSV do SYNC");
                }
            }
        }
        public static void ConvertErrosNationalPurchaseSync(Stream file, string nome)
        {
            if (file is not null)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int cols = worksheet.Dimension.End.Column;
                    int rows = worksheet.Dimension.End.Row;

                    var listToCsv = new List<NationalPurchaseSync>();

                    for (int i = 1; i <= rows; i++)
                    {
                        for (int j = 1; j <= cols; j++)
                        {
                            if (worksheet.Cells[i, j] != null && worksheet.Cells[i, j].Value != null)
                            {
                                var linha = "campo nao encontrado na linha A" + i;
                                var cellValue = worksheet.Cells[i, j].Value.ToString();

                                if (cellValue!.Contains("errorMessage"))
                                {
                                    var splitedValue = cellValue.Split("Message:");

                                    if (splitedValue.Length > 0)
                                    {
                                        var messageJson = splitedValue.Count() > 1 ? splitedValue[1] : splitedValue[0];
                                        var descricaoErro = ConvertFieldError(messageJson);
                                        var tipoErroMsgValue = ConvertFieldTipoErro(messageJson);
                                        var numeroNota = string.Empty;
                                        var tes = "497";

                                        if (nome.Equals("ProtheusNationalPurchaseSyncConsumer"))
                                        {
                                            numeroNota = ConvertFieldJustNumber(messageJson, "DocumentNumber", linha);
                                            tes = ConvertFieldJustNumber(messageJson, "Tes", linha);
                                        }
                                        else
                                        {
                                            numeroNota = ConvertFieldJustNumber(messageJson, "DocumentNumberInvoice", linha);
                                        }

                                        var idProduto = ConvertFieldJustNumber(messageJson, "OriginalPartNumber", linha);
                                        var condicaoPagamento = ConvertFieldJustNumber(messageJson, "PaymentTerms", linha);
                                        var fabric = ConvertFieldWithReplace(messageJson, "CompanyCode", linha);
                                        var idNotaScott = ConvertFieldJustNumber(messageJson, "ScottOperationID", linha);


                                        var filial = ConvertFieldJustNumber(messageJson, "BillingLocation", linha);
                                        var natureza = ConvertFieldJustNumber(messageJson, "NatureOfTransiction", linha);





                                        listToCsv.Add(new NationalPurchaseSync()
                                        {
                                            NumeroNota = numeroNota.Trim(),
                                            Fornecedor = fabric.Trim(),
                                            CondicaoPagamento = condicaoPagamento,
                                            IdeProdutoScott = idProduto,
                                            Filial = filial.Trim(),
                                            Error = descricaoErro.ToUTF8(),
                                            IdNotaScott = idNotaScott.Trim(),
                                            Natureza = natureza.Trim(),
                                            Tes = tes.Trim(),
                                            TipoErro = tipoErroMsgValue.Trim().ToUTF8()
                                        });

                                    }
                                }
                            }
                        }
                    }

                    var syncs = listToCsv.Distinct().ToList();

                    var dataStg = DateTime.Now.ToString();
                    dataStg = Regex.Replace(dataStg, "[^0-9a-zA-Z]+", "");

                    using (var writer = new StreamWriter(@"C:\Users\rogarcia\OneDrive - ScanSource, Inc\Documentos\ErrosLog\" + nome + dataStg + ".csv"))

                    using (var csv = new CsvWriter(writer, CultureInfo.CurrentCulture))
                    {
                        csv.WriteRecords(syncs);
                    }

                    Console.WriteLine("Fim da criação do CSV do " + nome);
                }
            }
        }


        public static string ToUTF8(this string text)
        {
            byte[] bytes = Encoding.GetEncoding(1252).GetBytes(text);
            var nameFixed = Encoding.UTF8.GetString(bytes);

            StringBuilder sbReturn = new StringBuilder();
            var arrayText = nameFixed.Normalize(NormalizationForm.FormD).ToCharArray();
            foreach (char letter in arrayText)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(letter) != UnicodeCategory.NonSpacingMark)
                    sbReturn.Append(letter);
            }
            return sbReturn.ToString();
        }

        public static string ConvertFieldTipoErro(string messageJson)
        {
            var errosMsgSplit = messageJson.Split("errorMessage\"\":\"\"\\\"\"");
            var errosMsgKey = string.Empty;
            var errosMsgValue = string.Empty;
            var descErro = string.Empty;
            var descricaoErro = string.Empty;

            if (messageJson.Equals("StatusCode: RequestTimeout,  "))
            {
                return "StatusCode: RequestTimeout";
            }

            if (messageJson.StartsWith("StatusCode: UnprocessableEntity,"))
            {
                return "StatusCode: 422";
            }

            if (messageJson.StartsWith("StatusCode: InternalServerError,  "))
            {
                return "StatusCode: 502";
            }

            if (messageJson.StartsWith("StatusCode: 0, "))
            {
                return "StatusCode: ,0";
            }

            errosMsgSplit = messageJson.Split("errorMessage\"\":\"\"\\\"\"");
            errosMsgKey = errosMsgSplit.Count() > 1 ? errosMsgSplit[1] : errosMsgSplit[0];

            if (errosMsgKey.Contains("nao foi localizado na filial"))
            {
                return "Produto nao localizado na filial";
            }

            errosMsgValue = errosMsgKey.Substring(0, errosMsgKey.IndexOf("\\\"\"\"\",\"\""));
            descErro = errosMsgValue.Replace("\\\\\\\"\"", "!");
            descricaoErro = descErro.Replace('!', '"');


            var listaErros = descricaoErro.Split("message\":\"");
            var tipoErrokey = string.Empty;
            var tipoErroMsgValue = string.Empty;

            if (listaErros.Count() > 1)
            {
                if (listaErros.Count() > 6)
                {
                    tipoErrokey = listaErros[6];
                }
                else if (listaErros.Count() == 2)
                {
                    tipoErrokey = listaErros[1];
                }
            }
            else
            {
                tipoErrokey = listaErros[0];
            }

            if (listaErros.Count() > 1 && tipoErrokey.Contains("An invalid response was received from the upstream serve"))
            {
                tipoErroMsgValue = tipoErrokey.Substring(0, 57);
            }
            else
            {
                tipoErroMsgValue = listaErros.Count() > 1 ?
   tipoErrokey.Substring(0, tipoErrokey.IndexOf("\",\""))
   : tipoErrokey;
            }

            return tipoErroMsgValue;
        }
        public static string ConvertFieldError(string messageJson)
        {
            var errosMsgValue = string.Empty;
            var errosMsgSplit = messageJson.Split("errorMessage\"\":\"\"\\\"\"");
            var errosMsgKey = errosMsgSplit.Count() > 1 ? errosMsgSplit[1]: errosMsgSplit[0];           

            if (errosMsgKey.Contains("nao foi localizado na filial"))
            {
                if (errosMsgSplit.Count() == 1)
                {
                    errosMsgSplit = messageJson.Split("errorMessage\\\"\":\\\"\"\\\\\\\"\"");
                    errosMsgKey = errosMsgSplit.Count() > 1 ? errosMsgSplit[1] : errosMsgSplit[0];
                }

                return errosMsgKey;
            }

             errosMsgValue = errosMsgKey.Substring(0, errosMsgKey.IndexOf("\\\"\"\"\",\"\""));
            var descErro = errosMsgValue.Replace("\\\\\\\"\"", "!");
            var descricaoErro = descErro.Replace('!', '"');

            if (messageJson.Contains("detailedMessage"))
            {
                errosMsgSplit = messageJson.Split("detailedMessage\\\\\\\"\":\\\\\\\"\"");
                errosMsgKey = errosMsgSplit[1];
                errosMsgValue = errosMsgKey.Substring(0, errosMsgKey.IndexOf("\\\"\"\"\",\"\""));
                descErro = errosMsgValue.Replace("\\\\\\\"\"", "!");
                descricaoErro = descErro.Replace('!', '"').Replace("\"\"}", "");
            }

            if (messageJson.Contains("Nota Fiscal jÃ¡ cadastrado com esse IDESCOT\\\\\\"))
            {
                errosMsgSplit = messageJson.Split("message\\\\\\\"\":\\\\\\\"\"");
                errosMsgKey = errosMsgSplit[1];
                errosMsgValue = errosMsgKey.Substring(0, errosMsgKey.IndexOf("\\\"\"\"\",\"\""));
                descErro = errosMsgValue.Replace("\\\\\\\"\"", "!");
                descricaoErro = descErro.Replace('!', '"').Replace("\"\"}", "").Replace("\",\"args\":[]}]}", "");
                var idNotaScott = ConvertFieldJustNumber(messageJson, "ScottOperationID", "");
                descricaoErro += " " + idNotaScott;

            }

            if (descricaoErro.Equals("Error Post: "))
            {
                descricaoErro = "Erro nao retornou no log";
            }


            if (descricaoErro.Contains("Codigo do fornecedor"))
            {
                descricaoErro = descricaoErro.Replace("6", "");
            }

            if (descricaoErro.Contains("Error Post:"))
            {
                descricaoErro = descricaoErro.Replace("Error Post:", "");
            }

            return descricaoErro;
        }
        public static string ConvertFieldWithReplace(string messageJson, string valueField, string linha)
        {
            var split = messageJson.Split(valueField);
            var Key = split.Count() > 1 ? split[1] : split[0];
            var value = Key.Substring(0, Key.IndexOf("\","));
            var valueStringFormat = messageJson.Contains(valueField)
                ? value.Replace('"', ' ').Replace("\\", "").Replace(':', ' ').Trim()
                : linha;

            return valueStringFormat;
        }
        public static string ConvertFieldJustNumber(string messageJson, string valueField, string linha)
        {
            var split = messageJson.Split(valueField);
            var key = split.Count() > 1 ? split[1] : split[0];
            var value = key.Substring(0, key.IndexOf("\","));
            var ret = Regex.Replace(value, "[^0-9a-zA-Z]+", "");

            return ret;
        }
        public static string ConvertFieldWithReplaceOriginalPartNumber(string messageJson, string valueField, string linha)
        {
            var stateSplit = messageJson.Split(valueField);
            var stateKey = stateSplit.Count() > 1 ? stateSplit[1] : stateSplit[0];
            var stateValue = stateKey.Substring(0, stateKey.IndexOf("\","));
            var partNumberOriginal = messageJson.Contains(valueField)
                ? stateValue.Replace('"', ' ').Replace(':', ' ').Trim()
                : linha;

            return partNumberOriginal;
        }

        public static string ConvertFieldModelo(string messageJson)
        {
            var stateSplit = messageJson.Split("B1_MODELO\"\":");
            var stateKey = stateSplit.Count() > 1 ? stateSplit[1] : stateSplit[0];
            var stateValue = stateKey.Substring(0, stateKey.IndexOf("\","));
            var partNumberOriginal = stateValue.Remove(0, 3).Replace('"', ' ').Trim();

            return partNumberOriginal;

        }
    }
}
