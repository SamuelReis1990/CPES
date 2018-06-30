using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using CPES.Models;
using static System.Environment;

namespace CPES
{
    public static class CPES
    {
        public static EnumerableRowCollection<DataRow> GetExcel(string caminhoArquivo, string tabela = "Sheet1$", int numColMatricula = 1, int numColNumero = 6, int numColCEP = 11)
        {
            var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", caminhoArquivo);
            var oleDbDataAdapter = new OleDbDataAdapter("SELECT * FROM [" + tabela + "]", connectionString);
            var dataSet = new DataSet();

            oleDbDataAdapter.Fill(dataSet, tabela);
            DataTable data = dataSet.Tables[tabela];

            DataTable dataTable = new DataTable();
            dataTable = data.Clone();
            dataTable.Columns[numColMatricula - 1].DataType = typeof(string);
            dataTable.Columns[numColNumero - 1].DataType = typeof(string);
            dataTable.Columns[numColCEP - 1].DataType = typeof(string);

            foreach (DataRow row in data.Rows)
            {
                dataTable.ImportRow(row);
            }

            return dataTable.AsEnumerable();
        }

        public static string[] GetDados(EnumerableRowCollection<DataRow> planilhaParametro, EnumerableRowCollection<DataRow> planilhaEndereco, EnumerableRowCollection<DataRow> planilhaLotacao, DataColumnCollection colParametro, DataColumnCollection colEndereco, DataColumnCollection colLotacao
                                        , int numColMatricula = 1
                                        , int numColMatriculaEndereco = 1
                                        , int numColMatriculaLotacao = 1
                                        , int numColNome = 2
                                        , int numColEndereco = 5
                                        , int numColNumero = 6
                                        , int numColComplemento = 7
                                        , int numColBairro = 8
                                        , int numColCidade = 9
                                        , int numColUF = 10
                                        , int numColCEP = 11
                                        , int numColLotacao = 26
                                        )
        {
            var dadosMatricula = planilhaParametro.Where(w => !String.IsNullOrWhiteSpace(w.Field<string>(colParametro[numColMatricula - 1].ToString())))
                .GroupBy(g => g.Field<string>(colParametro[numColMatricula - 1].ToString()))
                .Select(s => new DadosMatricula
                {
                    Matricula = s.Key
                }).OrderBy(o => o.Matricula).ToList();

            var dadosEndereco = planilhaEndereco.Where(w => !String.IsNullOrWhiteSpace(w.Field<string>(colEndereco[numColMatriculaEndereco - 1].ToString())) &&
                                                      planilhaParametro.Where(r => !String.IsNullOrWhiteSpace(r.Field<string>(colParametro[numColMatricula - 1].ToString())))
                                                      .Any(a => a.Field<string>(colParametro[numColMatricula - 1].ToString()) == w.Field<string>(colEndereco[numColMatriculaEndereco - 1].ToString())))
                    .Select(s => new DadosEndereco
                    {
                        Endereco = s.Field<string>(colEndereco[numColEndereco - 1].ToString()),
                        Numero = s.Field<string>(colEndereco[numColNumero - 1].ToString()),
                        Complemento = s.Field<string>(colEndereco[numColComplemento - 1].ToString()),
                        Bairro = s.Field<string>(colEndereco[numColBairro - 1].ToString()),
                        Cidade = s.Field<string>(colEndereco[numColCidade - 1].ToString()),
                        UF = s.Field<string>(colEndereco[numColUF - 1].ToString()),
                        CEP = s.Field<string>(colEndereco[numColCEP - 1].ToString())
                    }).ToList();

            var dadosLotacao = planilhaLotacao.Where(w => !String.IsNullOrWhiteSpace(w.Field<string>(colLotacao[numColMatriculaLotacao - 1].ToString())) &&
                                                     planilhaParametro.Where(r => !String.IsNullOrWhiteSpace(r.Field<string>(colParametro[numColMatricula - 1].ToString())))
                                                     .Any(a => a.Field<string>(colParametro[numColMatricula - 1].ToString()) == w.Field<string>(colLotacao[numColMatriculaLotacao - 1].ToString())))
                    .Select(s => new DadosLotacao
                    {
                        Nome = s.Field<string>(colLotacao[numColNome - 1].ToString()),
                        Lotacao = s.Field<string>(colLotacao[numColLotacao - 1].ToString())
                    }).ToList();

            var mergeMatEnd = dadosMatricula.Join(dadosEndereco, m => dadosMatricula.IndexOf(m), e => dadosEndereco.IndexOf(e), (m, e) => new { dadosEndereco = e, dadosMatricula = m }).ToList();

            var mergeListas = mergeMatEnd.Join(dadosLotacao, m => mergeMatEnd.IndexOf(m), l => dadosLotacao.IndexOf(l), (m, l) => new { dadosLotacao = l, dadosMergeMatriculaEndereco = m }).ToList();

            var dadosTrabalhador = mergeListas.Select(s => new DadosTrabalhador
            {
                Matricula = s.dadosMergeMatriculaEndereco.dadosMatricula.Matricula,
                Nome = s.dadosLotacao.Nome,
                Endereco = s.dadosMergeMatriculaEndereco.dadosEndereco.Endereco,
                Numero = s.dadosMergeMatriculaEndereco.dadosEndereco.Numero,
                Complemento = s.dadosMergeMatriculaEndereco.dadosEndereco.Complemento,
                Bairro = s.dadosMergeMatriculaEndereco.dadosEndereco.Bairro,
                Cidade = s.dadosMergeMatriculaEndereco.dadosEndereco.Cidade,
                UF = s.dadosMergeMatriculaEndereco.dadosEndereco.UF,
                CEP = s.dadosMergeMatriculaEndereco.dadosEndereco.CEP,
                Lotacao = s.dadosLotacao.Lotacao
            }).ToList();

            string[] retorno = CreateCsvSucesso(dadosTrabalhador);
            retorno[2] = dadosTrabalhador.Count().ToString();

            return retorno;
        }

        public static string[] CreateCsvSucesso(List<DadosTrabalhador> dadosTrabalhador)
        {
            string nomeArquivo = "CPES_" + DateTime.Now.ToString("ddMMyyyyHHmmss") + ".csv";            
            string caminho = GetFolderPath(SpecialFolder.DesktopDirectory);
            FileStream file = new FileStream(caminho + "\\" + nomeArquivo, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            StreamWriter writer = new StreamWriter(file, Encoding.Default);

            writer.WriteLine("matricula;nome;endereco;lotacao".ToUpper());

            foreach (var dados in dadosTrabalhador)
            {
                string texto = dados.Matricula.ToString() + ";" +
                               dados.Nome + ";" +
                               dados.Endereco + ", " + dados.Numero + " - " + dados.Complemento + " - " + dados.Bairro + " - " + dados.Cidade + " - " + dados.UF + " - " + dados.CEP + ";" +
                               dados.Lotacao;

                writer.WriteLine(texto.ToUpper());
            }

            writer.Flush();

            return new string[] { nomeArquivo, caminho, null };
        }

        public static string[] CreateCsvErro(Erro erro)
        {
            string nomeArquivo = "log_erro_" + DateTime.Now.ToString("ddMMyyyy") + ".csv";
            string caminho = GetFolderPath(SpecialFolder.DesktopDirectory);
            FileStream file;
            StreamWriter writer;

            if (!File.Exists(caminho + "\\" + nomeArquivo))
            {
                file = new FileStream(caminho + "\\" + nomeArquivo, FileMode.CreateNew, FileAccess.ReadWrite);
                writer = new StreamWriter(file, Encoding.Default);
                writer.WriteLine("Data;Mensagem;StackTrace;CaminhoArquivoParametro;CaminhoArquivoEndereco;CaminhoArquivoLotacao;TabelaParametro;TabelaEndereco;TabelaLotacao;NumColMatricula;NumColMatriculaEndereco;NumColMatriculaLotacao;NumColNome;NumColEndereco;NumColNumero;NumColComplemento;NumColBairro;NumColCidade;NumColUF;NumColCEP;NumColLotacao");
            }
            else
            {
                file = new FileStream(caminho + "\\" + nomeArquivo, FileMode.Append);
                writer = new StreamWriter(file, Encoding.Default);
            }

            string texto = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + ";" +
                           erro.Mensagem + ";" +
                           erro.StackTrace + ";" +
                           erro.CaminhoArquivoParametro + ";" +
                           erro.CaminhoArquivoEndereco + ";" +
                           erro.CaminhoArquivoLotacao + ";" +
                           erro.TabelaParametro + ";" +
                           erro.TabelaEndereco + ";" +
                           erro.TabelaLotacao + ";" +
                           erro.NumColMatricula + ";" +
                           erro.NumColMatriculaEndereco + ";" +
                           erro.NumColMatriculaLotacao + ";" +
                           erro.NumColNome + ";" +
                           erro.NumColEndereco + ";" +
                           erro.NumColNumero + ";" +
                           erro.NumColComplemento + ";" +
                           erro.NumColBairro + ";" +
                           erro.NumColCidade + ";" +
                           erro.NumColUF + ";" +
                           erro.NumColCEP + ";" +
                           erro.NumColLotacao;

            writer.WriteLine(texto);
            writer.Flush();

            return new string[] { nomeArquivo, caminho };
        }
    }
}
