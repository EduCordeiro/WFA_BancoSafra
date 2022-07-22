
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Compression;
using System.Diagnostics;
using System.Threading;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        private static int Lote = 0;
        private static int SequencialFAC = 0;
        private static DateTime DataPostagem;
        private static string ArquivoMDB = AppDomain.CurrentDomain.BaseDirectory + @"Dados\Dados.MDB";
        private static string ArquivoMidias = AppDomain.CurrentDomain.BaseDirectory + @"Midias\";
        private static List<Triagem> ListaTriagem = new List<Triagem>();
        private static Relatorio RelatorioFinal = new Relatorio();

        private static string drpostagem = "72".PadLeft(2, '0');
        private static string codigoadmistrativo = "11181591".PadLeft(8, '0');
        private static string contratofac = "9912279437 ".PadLeft(10, '0');
        private static string cartaopostagem = "0070596450".PadLeft(12, '0');
        private static string cartaopostagem02 = "0072162848".PadLeft(12, '0');
        private static string cartaopostagem03 = "0069385335".PadLeft(12, '0');

        private static string codigopostagem = "425791".PadLeft(8, '0');
        private static string ceppostagem = "05311900".PadLeft(8, '0');

        private static string cliente = "BANCO SAFRA";
        private static string produto = "";

        private static string NM_ARQUIVO = "";
        private static string DataMovimento = "";

        private static List<string> ListaRelatorio = new List<string>();
        private static List<string> ListaRetido = new List<string>();

        int txtIndice = 0;

        FileStream streamLogProc;
        StreamWriter swLogProc;

        public Form1()
        {
            InitializeComponent();
        }

        private void sairToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private static bool Processar(string Arquivo) // processa carne
        {
            try
            {
                NM_ARQUIVO = Path.GetFileNameWithoutExtension(Arquivo);
                Directory.CreateDirectory(Path.GetDirectoryName(Arquivo) + "\\Processado\\");

                OrdenarArquivo(Arquivo);
                Arquivo = Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo);
                var nrlote = 0;

                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                SeparaProdutos(Arquivo);

                ProcessarCarne(Path.ChangeExtension(Arquivo, "001"), "CDC");


                ProcessarCarne(Path.ChangeExtension(Arquivo, "058"), "CDC");
                ProcessarCarne(Path.ChangeExtension(Arquivo, "MUTUO001"), "MUTUO");
                ProcessarCarne(Path.ChangeExtension(Arquivo, "MUTUO058"), "MUTUO");

                ProcessarCarne(Path.ChangeExtension(Arquivo, "MUTUO_PF001"), "MUTUO_PF");
                ProcessarCarne(Path.ChangeExtension(Arquivo, "MUTUO_PF058"), "MUTUO_PF");



                SepararImpressao(Path.ChangeExtension(Arquivo, "001TMP"));
                SepararImpressao(Path.ChangeExtension(Arquivo, "058TMP"));
                SepararImpressao(Path.ChangeExtension(Arquivo, "MUTUO001TMP"));
                SepararImpressao(Path.ChangeExtension(Arquivo, "MUTUO058TMP"));

                SepararImpressao(Path.ChangeExtension(Arquivo, "MUTUO_PF001TMP"));
                SepararImpressao(Path.ChangeExtension(Arquivo, "MUTUO_PF058TMP"));

                StreamWriter RelatorioProcessamento = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Relatorio_Processamento_" + NM_ARQUIVO + "_" + string.Format("{0:ddMMyyyy}", DateTime.Now) + ".TXT", false, Encoding.GetEncoding("ISO-8859-1"));

                RelatorioProcessamento.WriteLine("Resumo do Processamento:");
                RelatorioProcessamento.WriteLine("Data de Postagem:" + string.Format("{0:dd/MM/yyyy}", DataPostagem));
                RelatorioProcessamento.WriteLine("");
                RelatorioProcessamento.WriteLine("");

                for (int i = 0; i < 100; i++)
                {
                    if (RelatorioFinal.Carnes[i] > 0)
                        RelatorioProcessamento.WriteLine(RelatorioFinal.Carnes[i].ToString().PadLeft(3) + " carnes de " + i.ToString().PadLeft(2) + " laminas");
                }

                RelatorioProcessamento.WriteLine("");
                RelatorioProcessamento.WriteLine("");
                RelatorioProcessamento.WriteLine("");

                int iLamVal = 0, iLamMsg = 0, iLamExt = 0;
                int iTotCrV = 0, iCarBco = 0, iTotLam = 0;


                iTotCrV = RelatorioFinal.CarnesValidos;
                iCarBco = RelatorioFinal.CarnesBranco;

                iLamVal = RelatorioFinal.LaminasValidas;
                iLamMsg = RelatorioFinal.LaminasMensagem + iTotCrV;
                iLamExt = RelatorioFinal.LaminasExtra + (iCarBco * 15);


                iTotLam = iLamVal + iLamMsg + iLamExt;

                RelatorioProcessamento.WriteLine("Laminas Validas:  " + RelatorioFinal.LaminasValidas.ToString().PadLeft(10));
                // RelatorioProcessamento.WriteLine("Laminas Mensagem: " + RelatorioFinal.LaminasMensagem.ToString().PadLeft(10));
                // RelatorioProcessamento.WriteLine("Laminas Extra:    " + RelatorioFinal.LaminasExtra.ToString().PadLeft(10));
                // RelatorioProcessamento.WriteLine("Total de Laminas: " + (RelatorioFinal.LaminasValidas + RelatorioFinal.LaminasMensagem + RelatorioFinal.LaminasExtra).ToString().PadLeft(10));

                RelatorioProcessamento.WriteLine("Laminas Mensagem: " + iLamMsg.ToString().PadLeft(10));
                RelatorioProcessamento.WriteLine("Laminas Extra:    " + iLamExt.ToString().PadLeft(10));
                RelatorioProcessamento.WriteLine("Total de Laminas: " + iTotLam.ToString().PadLeft(10));


                RelatorioProcessamento.WriteLine("");
                RelatorioProcessamento.WriteLine("");
                RelatorioProcessamento.WriteLine("");

                RelatorioProcessamento.WriteLine("Carnes Validos:   " + RelatorioFinal.CarnesValidos.ToString().PadLeft(10));
                RelatorioProcessamento.WriteLine("Carnes Branco:    " + RelatorioFinal.CarnesBranco.ToString().PadLeft(10));
                RelatorioProcessamento.WriteLine("----------------------------");
                RelatorioProcessamento.WriteLine("Total de Carnes:  " + (RelatorioFinal.CarnesValidos + RelatorioFinal.CarnesBranco).ToString().PadLeft(10));
                RelatorioProcessamento.Dispose();

                #region Grava Relatorio Manuseio
                int Pagina = 0;
                int Linha = 0;
                StreamWriter RelatorioManuseio = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Relatorio_Manuseio_" + NM_ARQUIVO + "_" + string.Format("{0:ddMMyyyy}", DateTime.Now) + ".TXT", false, Encoding.GetEncoding("ISO-8859-1"));

                foreach (var linha in ListaRelatorio)
                {
                    if (Linha % 34 == 0)
                    {
                        if (Pagina > 0)
                        {
                            RelatorioManuseio.WriteLine("".PadRight(170));
                            RelatorioManuseio.WriteLine("".PadRight(170));
                            RelatorioManuseio.WriteLine("".PadRight(170));
                            RelatorioManuseio.WriteLine("".PadRight(170));
                        }

                        Pagina++;
                        RelatorioManuseio.WriteLine("BANCO SAFRA                       RELACAO  DE  BLOQUETES  EMITIDOS                             DT.EMIS: " + string.Format("{0:dd/MM/yyyy}", DateTime.Now) + "                                           pag." + Pagina.ToString("d4"));
                        RelatorioManuseio.WriteLine("                                            V E I C U L O S                                    DT.BASE: " + string.Format("{0:dd/MM/yyyy}", DateTime.Now) + "                                                    ");
                        RelatorioManuseio.WriteLine("".PadRight(166));
                        RelatorioManuseio.WriteLine("CONTRATO         PLANO  1.VENCTO  REVENDA  NOME                                       CIDADE                        BAIRRO                        CEP       UF  ENDERECO                                                                                            AUDIT   ");
                        RelatorioManuseio.WriteLine("----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
                    }

                    Linha++;
                    RelatorioManuseio.WriteLine(linha);
                }
                RelatorioManuseio.WriteLine("".PadRight(170));
                RelatorioManuseio.WriteLine("".PadRight(170));
                RelatorioManuseio.WriteLine("".PadRight(170));
                RelatorioManuseio.WriteLine("TOTAL DE CONTRATOS: " + ListaRelatorio.Count());

                if (ListaRetido.Count > 0)
                {
                    RelatorioManuseio.WriteLine("".PadRight(170));
                    RelatorioManuseio.WriteLine("".PadRight(170));
                    RelatorioManuseio.WriteLine("Contratos/Parcelas não impressos".PadRight(166));

                    RelatorioManuseio.WriteLine("".PadRight(170));
                    RelatorioManuseio.WriteLine("CONTRATO         PLANO  1.VENCTO  REVENDA  NOME                                       CIDADE           BAIRRO           CEP       UF  ENDERECO                            ");
                    RelatorioManuseio.WriteLine("--------------------------------------------------------------------------------------------------------------------------------------------------------------------------");

                    foreach (var linha in ListaRetido)
                    {
                        RelatorioManuseio.WriteLine(linha);
                    }
                    RelatorioManuseio.WriteLine("".PadRight(170));
                    RelatorioManuseio.WriteLine("".PadRight(170));
                    RelatorioManuseio.WriteLine("".PadRight(170));
                    RelatorioManuseio.WriteLine("TOTAL DE PARCELAS RETIDAS: " + ListaRetido.Count());
                }

                RelatorioManuseio.Dispose();
                #endregion
                return true;
            }
            catch
            {
                return false;
            }

        }

        private static void OrdenarArquivo(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();
                string CarneOld = "";

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));

                while (!sr.EndOfStream)
                {
                    string linha = sr.ReadLine();

                    try
                    {
                        #region
                        if (linha.Substring(000, 001) == "1")
                        {
                            sw.WriteLine(linha);
                        }

                        #region Tipo 2
                        else if (linha.Substring(000, 001) == "2")
                        {
                            if (CarneOld != linha.Substring(006, 016).Trim())
                            {
                                if (!string.IsNullOrEmpty(CarneOld))
                                {
                                    _Carnes.Add(_CarneAtual);
                                    _CarneAtual = new CarneOrdenaCEP();
                                }
                                CarneOld = linha.Substring(006, 016).Trim();
                                _CarneAtual.DadosCarne = CarneOld;
                                _CarneAtual.CEP = Convert.ToInt32(linha.Substring(856, 008).Trim().PadLeft(8, '0'));
                            }
                            _CarneAtual.Parcelas.Add(linha);
                        }
                        #endregion

                        #region Tipo 4
                        else if (linha.Substring(000, 001) == "4")
                        {
                            #region Ultimo Carne
                            _Carnes.Add(_CarneAtual);
                            #endregion

                            var lista = _Carnes.OrderBy(x => x.CEP);
                            foreach (var item in lista)
                            {
                                foreach (var parcela in item.Parcelas)
                                    sw.WriteLine(parcela);
                            }

                            //Trailler
                            sw.WriteLine(linha);
                            sw.Dispose();
                        }
                        #endregion
                        else
                            _CarneAtual.Parcelas.Add(linha);
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(linha);
                    }
                }

            }
        }

        private static void OrdenarCartaCQC(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();
                sw.WriteLine(linha);

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');
                    Int32 CEP = 0;
                    Int32.TryParse(tmp[6].Replace("-", "").Trim() == "" ? "95096100" : tmp[6].Replace("-", "").Trim(), out CEP);

                    _CarneAtual.CEP = CEP;
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new CarneOrdenaCEP();

                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }

        private static void OrdenarCarta(string Arquivo, string path = "")
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + path + "\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();
                sw.WriteLine(linha);

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');
                    Int32 CEP = 0;
                    Int32.TryParse(tmp[6].Replace("-", "").Trim() == "" ? "95096100" : tmp[6].Replace("-", "").Trim(), out CEP);

                    _CarneAtual.CEP = CEP;
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new CarneOrdenaCEP();

                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }

        private static void SeparaProdutos(string Arquivo)
        {
            File.Delete(Path.GetDirectoryName(Arquivo) + "//Relatorio_" + NM_ARQUIVO + "_" + string.Format("{0:ddMMyyyy}", DateTime.Now) + ".TXT");
            StreamWriter Relatorio = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Relatorio_" + NM_ARQUIVO + "_" + string.Format("{0:ddMMyyyy}", DateTime.Now) + ".TXT", true, Encoding.GetEncoding("ISO-8859-1"));
            Relatorio.WriteLine("Data da Operação" + ";" +
                                "Contrato" + ";" +
                                "CPF_CNPJ" + ";" +
                                "Chassi" + ";" +
                                "Quantidade de parcelas" + ";" +
                                "Produto" + ";" +
                                "Data da Impressão do carnê" + ";" +
                                "Data de envio da postagem" + ";" +
                                "CEP;");
            Relatorio.Dispose();

            try
            {
                StreamWriter Arquivo01 = new StreamWriter(Path.ChangeExtension(Arquivo, "001"), false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter Arquivo58 = new StreamWriter(Path.ChangeExtension(Arquivo, "058"), false, Encoding.GetEncoding("ISO-8859-1"));

                StreamWriter Arquivo01M = new StreamWriter(Path.ChangeExtension(Arquivo, "MUTUO001"), false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter Arquivo58M = new StreamWriter(Path.ChangeExtension(Arquivo, "MUTUO058"), false, Encoding.GetEncoding("ISO-8859-1"));

                StreamWriter Arquivo01MPF = new StreamWriter(Path.ChangeExtension(Arquivo, "MUTUO_PF001"), false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter Arquivo58MPF = new StreamWriter(Path.ChangeExtension(Arquivo, "MUTUO_PF058"), false, Encoding.GetEncoding("ISO-8859-1"));

                bool GravaCarne = true;

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    string tipoarquivo = "";
                    string tipocarne = "";

                    while (!sr.EndOfStream)
                    {
                        string linha = sr.ReadLine();

                        if (linha.Substring(00, 01) == "1")
                            DataMovimento = linha.Substring(007, 002) + "/" + linha.Substring(005, 002) + "/" + linha.Substring(001, 004);

                        if (linha.Substring(00, 01) == "1" || linha.Substring(00, 01) == "2" || linha.Substring(00, 01) == "4")
                            GravaCarne = true;

                        if (linha.Substring(00, 01) == "2")
                        {
                            //if (linha.Substring(614, 001).ToUpper() == "N" || linha.Substring(623, 001).ToUpper() == "N")
                            if (linha.Substring(623, 001).ToUpper() == "N")
                            {
                                GravaCarne = false;
                                ListaRetido.Add(linha.Substring(006, 016).PadRight(020) + //CONTRATO
                                                linha.Substring(486, 003).PadRight(004) + //PLANO
                                               (linha.Substring(039, 002) + "/" +
                                                linha.Substring(037, 002) + "/" +
                                                linha.Substring(033, 004)).PadRight(011) + //1.VENCTO
                                                linha.Substring(519, 006).PadRight(008) + //REVENDA
                                                linha.Substring(087, 030).PadRight(043) + //NOME
                                                linha.Substring(824, 030).PadRight(017) + //CIDADE
                                                linha.Substring(794, 030).PadRight(030) + //BAIRRO
                                                linha.Substring(856, 008).PadRight(010) + //CEP
                                                linha.Substring(854, 002).PadRight(004) + //UF
                                               (linha.Substring(699, 060).Trim() + " " +
                                                linha.Substring(759, 035).Trim()).PadRight(100) //ENDERECO
                                               );
                            }
                        }

                        if (GravaCarne)
                        {
                            if (linha.Substring(00, 01) == "1" || linha.Substring(00, 01) == "4")
                            {
                                Arquivo01.WriteLine(linha);
                                Arquivo58.WriteLine(linha);
                                Arquivo01M.WriteLine(linha);
                                Arquivo58M.WriteLine(linha);

                                Arquivo01MPF.WriteLine(linha);
                                Arquivo58MPF.WriteLine(linha);
                            }
                            else
                            {
                                if (linha.Substring(00, 01) == "2")
                                {
                                    tipoarquivo = linha.Substring(684, 002);
                                    tipocarne = linha.Substring(614, 001);
                                }


                                if (tipocarne == "S")
                                {
                                    if (tipoarquivo == "01")
                                        Arquivo01.WriteLine(linha);
                                    else if (tipoarquivo == "58")
                                        Arquivo58.WriteLine(linha);
                                }


                                if (tipocarne == "N")
                                {
                                    if (tipoarquivo == "01")
                                        Arquivo01.WriteLine(linha);
                                    else if (tipoarquivo == "58")
                                        Arquivo58.WriteLine(linha);
                                }

                                if (tipocarne == "M")
                                {
                                    if (tipoarquivo == "01")
                                        Arquivo01M.WriteLine(linha.PadRight(900, ' '));
                                    else if (tipoarquivo == "58")
                                        Arquivo58M.WriteLine(linha.PadRight(900, ' '));
                                }

                                if (tipocarne == "F")
                                {
                                    if (tipoarquivo == "01")
                                        Arquivo01MPF.WriteLine(linha.PadRight(900, ' '));
                                    else if (tipoarquivo == "58")
                                        Arquivo58MPF.WriteLine(linha.PadRight(900, ' '));
                                }



                            }
                        }
                    }
                }

                Arquivo01.Dispose();
                Arquivo58.Dispose();
                Arquivo01M.Dispose();
                Arquivo58M.Dispose();

                Arquivo01MPF.Dispose();
                Arquivo58MPF.Dispose();

            }
            catch
            { }
        }

        private static void SepararImpressao(string Arquivo)
        {
            try
            {
                StreamWriter       Capa = new StreamWriter(Arquivo.Replace(".001TMP", "_001_CAPA.TXT"      ).Replace(".058TMP", "_058_CAPA.TXT"      ).Replace(".MUTUO001TMP", "MUTUO_001_CAPA.TXT"      ).Replace(".MUTUO058TMP", "MUTUO_058_CAPA.TXT"      ).Replace(".MUTUO_PF058TMP", "MUTUO_PF_058_CAPA.TXT"), false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter      Miolo = new StreamWriter(Arquivo.Replace(".001TMP", "_001_MIOLO.TXT"     ).Replace(".058TMP", "_058_MIOLO.TXT"     ).Replace(".MUTUO001TMP", "MUTUO_001_MIOLO.TXT"     ).Replace(".MUTUO058TMP", "MUTUO_058_MIOLO.TXT"     ).Replace(".MUTUO_PF058TMP", "MUTUO_PF_058_MIOLO.TXT"), false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter ContraCapa = new StreamWriter(Arquivo.Replace(".001TMP", "_001_CONTRACAPA.TXT").Replace(".058TMP", "_058_CONTRACAPA.TXT").Replace(".MUTUO001TMP", "MUTUO_001_CONTRACAPA.TXT").Replace(".MUTUO058TMP", "MUTUO_058_CONTRACAPA.TXT").Replace(".MUTUO_PF058TMP", "MUTUO_PF_058_CONTRACAPA.TXT"), false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    while (!sr.EndOfStream)
                    {
                        var linha = sr.ReadLine();
                        if (linha.Substring(008, 002) == "09" || linha.Substring(008, 002) == "08" || linha.Substring(008, 002) == "10" || linha.Substring(008, 002) == "11" || linha.Substring(008, 002) == "18")
                            Capa.WriteLine(linha);
                        else if (linha.Substring(008, 002) == "01")
                            ContraCapa.WriteLine(linha);
                        else
                            Miolo.WriteLine(linha);
                    }
                }
                Capa.Dispose();
                Miolo.Dispose();
                ContraCapa.Dispose();
                File.Delete(Arquivo);
            }
            catch
            { }
        }

        private static void ProcessarCarne(string Arquivo, string Tipocarne)
        {
            int ncnt = 0;
            int nponto = 0;
            try
            {
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                StreamWriter Relatorio = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Relatorio_" + NM_ARQUIVO + "_" + string.Format("{0:ddMMyyyy}", DateTime.Now) + ".TXT", true, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    List<Carne> _Carnes = new List<Carne>();
                    Carne _CarneAtual = new Carne();
                    string CarneOld = "";
                    string Lamina = "";
                    string Empresa = "";
                    int SeqCarne = 0;

                   
                    while (!sr.EndOfStream)
                    {
                        
                              ncnt++;

                        if (ncnt == 72680) 
                        {
                           ///  MessageBox.Show("**** parou aqui ...");
                        }

                        string linha = sr.ReadLine();

                        #region Tipo 2
                        if (linha.Substring(000, 001) == "2")
                        {
                            if (CarneOld != linha.Substring(006, 016).Trim())
                            {
                                if (!string.IsNullOrEmpty(CarneOld))
                                {
                                    if (_CarneAtual.Parcelas.Count > 0)
                                    {
                                        if (_CarneAtual.Parcelas[0].Substring(614, 001).ToUpper() != "N")
                                        {
                                            _CarneAtual.Qtdeparcela++;
                                        }
                                    }

                                    _Carnes.Add(_CarneAtual);
                                    _CarneAtual = new Carne();
                                }
                                SeqCarne++;
                                CarneOld = linha.Substring(006, 016).Trim();
                                _CarneAtual.DadosCarne = CarneOld;
                                _CarneAtual.Seqcarne = SeqCarne;
                            }
                            Lamina = linha;
                            Empresa = linha.Substring(684, 002);
                        }
                        #endregion

                        #region Tipo 5
                        else if (linha.Substring(000, 001) == "5")
                        {
                            var valor = Convert.ToDouble(Lamina.Substring(236, 015)) / 100;
                            var barra = Regex.Replace(Lamina.Substring(289, 057), "[^0-9]", "");
                            try
                            {
                                barra = barra.Substring(00, 04) + barra.Substring(32, 15) +
                                        barra.Substring(04, 05) + barra.Substring(10, 10) +
                                        barra.Substring(21, 10);
                            }
                            catch
                            {
                                barra = "";
                            }

                            var CIF = "8888888888888888888888888888888888";
                            var CDD = "CDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDC";

                            _CarneAtual.Parcelas.Add(Lamina + linha + string.Format("{0:0,0.00}", valor).PadLeft(15) + string.Format("{0:dd/MM/yyyy}", DateTime.Now).PadRight(10) + barra.PadRight(60) + CIF.PadRight(34) + CDD.PadRight(100) + Path.GetFileName(Arquivo).PadRight(30));
                            _CarneAtual.Qtdeparcela++;
                        }
                        #endregion

                        #region Tipo 3
                        else if (linha.Substring(000, 001) == "3")
                        {
                            _CarneAtual.Parcelas[0] = _CarneAtual.Parcelas[0] + linha;
                        }
                        #endregion

                        #region Tipo 4
                        else if (linha.Substring(000, 001) == "4")
                        {
                            #region Ultimo Carne
                            if (_CarneAtual.Parcelas.Count > 0)
                            {
                                if (_CarneAtual.Parcelas[0].Substring(614, 001).ToUpper() != "N")
                                {
                                    _CarneAtual.Qtdeparcela++;
                                }
                            }
                            _Carnes.Add(_CarneAtual);
                            _CarneAtual = new Carne();
                            SeqCarne++;
                            CarneOld = linha.Substring(006, 016).Trim();
                            _CarneAtual.DadosCarne = CarneOld;
                            _CarneAtual.Seqcarne = SeqCarne;
                            #endregion

                            StreamWriter sw = new StreamWriter(Arquivo + "TMP", false, Encoding.GetEncoding("ISO-8859-1"));

                            var lista = _Carnes.OrderBy(x => x.Qtdeparcela);
                            List<Parcelas>[] lsttmp = new List<Parcelas>[2];
                            lsttmp[0] = new List<Parcelas>();
                            lsttmp[1] = new List<Parcelas>();
                            int seq = 0;
                            int qtdeparcela = 0;

                            foreach (var item in lista)
                            {
                                if (qtdeparcela > 0)
                                {
                                    if (seq == 2 || qtdeparcela != item.Qtdeparcela)
                                    {
                                        qtdeparcela = 0;
                                        int prc = lsttmp[0].Count;
                                        if (seq == 1)
                                        {
                                            if (prc > 0)
                                            {
                                                RelatorioFinal.CarnesBranco++;
                                                RelatorioFinal.LaminasExtra += prc + 7;
                                            }

                                            for (int j = 0; j < prc; j++)
                                            {
                                                Parcelas Parcela = new Parcelas();
                                                Parcela.DadosParcelas.Add(lsttmp[0][j].DadosParcelas[0].Substring(00, 20) + "*".PadRight(100, '*'));
                                                lsttmp[1].Add(Parcela);
                                            }
                                        }

                                        int jj=0, ii=0;
                                        try
                                        {
                                           
                                            for (int j = 0; j < prc; j++)
                                            {
                                                jj = j;
                                                for (int i = 0; i < 2; i++)
                                                {
                                                    ii = i;
                                                    sw.WriteLine(lsttmp[i][j].DadosParcelas[0]);
                                                }
                                            }
                                        }
                                        catch (  Exception   err )
                                        {

                                            MessageBox.Show("teste aqui ... " + " item="+ item +"    -   " + prc.ToString()+" j="+jj.ToString() + " i=" + ii.ToString());
                                            int n1 = jj;

                                        }


                                        seq = 0;
                                        lsttmp[0].Clear();
                                        lsttmp[1].Clear();
                                    }
                                }

                                int laminas = 0;
                                foreach (var parcela in item.Parcelas)
                                {
                                    laminas++;
                                    Parcelas Parcela = new Parcelas();
                                    if (lsttmp[seq].Count == 0)
                                    {
                                        SequencialFAC++;

                                        item.Seqcarne = SequencialFAC;

                                        Int32 CEP = 0;
                                        Int32.TryParse(parcela.Substring(856, 008).Trim() == "" ? "95096100" : parcela.Substring(856, 008).Trim(), out CEP);

                                        string Destino = "3";
                                        if (CEP < 10000000)
                                            Destino = "1";
                                        else if (CEP < 20000000)
                                            Destino = "2";

                                        var nmCDD = "Não Localizado";
                                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                                        if (CDD != null)
                                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                                        RelatorioFinal.CarnesValidos++;
                                        RelatorioFinal.LaminasValidas += item.Parcelas.Count + 2;
                                        RelatorioFinal.LaminasMensagem += 3;
                                        RelatorioFinal.Carnes[item.Parcelas.Count]++;

                                        #region Relatotio Manuseio
                                        ListaRelatorio.Add(parcela.Substring(006, 016).PadRight(020) + //CONTRATO
                                                           parcela.Substring(486, 003).PadRight(004) + //PLANO
                                                          (parcela.Substring(039, 002) + "/" +
                                                           parcela.Substring(037, 002) + "/" +
                                                           parcela.Substring(033, 004)).PadRight(011) + //1.VENCTO
                                                           parcela.Substring(519, 006).PadRight(008) + //REVENDA
                                                           parcela.Substring(087, 030).PadRight(043) + //NOME
                                                           parcela.Substring(824, 030).PadRight(017) + //CIDADE
                                                           parcela.Substring(794, 030).PadRight(030) + //BAIRRO
                                                           parcela.Substring(856, 008).PadRight(010) + //CEP
                                                           parcela.Substring(854, 002).PadRight(004) + //UF
                                                          (parcela.Substring(699, 060).Trim() + " " +
                                                           parcela.Substring(759, 035).Trim()).PadRight(100) + //ENDERECO
                                                           item.Seqcarne.ToString("d8")
                                                          );
                                        #endregion

                                        Relatorio.WriteLine(parcela.Substring(047, 002).Trim() + "/" +
                                                            parcela.Substring(045, 002).Trim() + "/" +
                                                            parcela.Substring(041, 004).Trim() + ";" + //Data da Operação
                                                            parcela.Substring(006, 016).Trim() + ";" + //Contrato
                                                            parcela.Substring(580, 014).Trim() + ";" + //CPF / CNPJ
                                                            parcela.PadRight(2200).Substring(2144, 020).Trim() + ";" + //Chassi
                                                            item.Parcelas.Count.ToString() + ";" + //Quantidade de parcelas
                                                            parcela.Substring(684, 002).Trim() + ";" + //Produto
                                                            string.Format("{0:dd/MM/yyyy}", DateTime.Now) + ";" + //Data da Impressão do carnê
                                                            string.Format("{0:dd/MM/yyyy}", DataPostagem) + ";" +  //Data de envio da postagem
                                                            parcela.Substring(856, 008).Trim() + ";"); //CEP

                                        Parcela.DadosParcelas.Add(item.Seqcarne.ToString("d8") + "09" + item.Qtdeparcela.ToString("d3") + parcela.Replace("8888888888888888888888888888888888", CIF).Replace("CDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDCDDC", nmCDD.PadRight(100)));
                                        
                                        carga.WriteLine(parcela.Substring(006, 016).Trim() + ";" + //chavepesquisa
                                                        parcela.Substring(087, 030).Trim() + ";" + //destinatario
                                                        parcela.Substring(087, 030).Trim() + ";" + //nomekit
                                                        parcela.Substring(087, 030).Trim() + ";" + //nome
                                                        parcela.Substring(699, 060).Trim() + ";" + //endereco
                                                        parcela.Substring(759, 035).Trim() + ";" + //numero
                                                        parcela.Substring(794, 030).Trim() + ";" + //complemento
                                                        "" + ";" + //bairro
                                                        parcela.Substring(824, 030).Trim() + ";" + //cidade
                                                        parcela.Substring(854, 002).Trim() + ";" + //estado
                                                        (parcela.Substring(856, 008).Trim() == "" ? "95096100" : parcela.Substring(856, 008).Trim()) + ";" + //cep
                                                        parcela.Substring(580, 014).Trim() + ";" + //cpfcnpj
                                                        CIF + ";" + //rastreio
                                                        CIF.Substring(26, 01) + ";" + //localidade
                                                        "FAC" + ";" + //tipopostagem
                                                        "Processado" + ";" + //status
                                                        "" + ";" + //referencia
                                                        Lote.ToString() + ";" + //os
                                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                                        SequencialFAC.ToString("d8") + ";" + //audit
                                                        (((item.Parcelas.Count + 2) * 2.19) + 15.8).ToString() + ";" + //peso
                                                        "1" + ";" + //paginainicial
                                                        "2" + ";" + //paginafinal
                                                        item.Parcelas.Count.ToString() + ";" + //qtdepaginas
                                                        item.Parcelas.Count.ToString() + ";" + //qtdefolhas
                                                        "P" + ";" + //familia
                                                        "" + ";" + //codinsumo01
                                                        "" + ";" + //descricao01
                                                        "" + ";" + //codinsumo02
                                                        "" + ";" + //descricao02
                                                        "" + ";" + //codinsumo03
                                                        "" + ";" + //descricao03
                                                        "" + ";" + //codinsumo04
                                                        "" + ";" + //descricao04
                                                        "" + ";" + //codinsumo05
                                                        "" + ";" + //descricao05
                                                        "" + ";" + //crt
                                                        "" + ";" + //etq
                                                        "" + ";" + //crn
                                                        "" + ";" + //manual
                                                        "" + ";" + //codigomanual
                                                        "" + ";" + //auditpostagem
                                                        "" + ";" + //auditkit
                                                        "" + ";" + //auditado
                                                        "" + ";" + //fechado
                                                        "" + ";" + //arquivoentradaid
                                                        "" + ";" + //datarecepcao
                                                        "" + ";" + //datapostagem
                                                        "" + ";" + //dataentrega
                                                        "" + ";" + //tipokit
                                                        "" + ";" + //leitura
                                                        "" + ";" + //email
                                                        "" + ";" + //contrato
                                                        "" + ";" + //nomepdf
                                                        "" + ";" + //telefone
                                                        "" + ";" + //pathsms
                                                        "" + ";" + //pathemail
                                                        "" + ";" + //pdfgerado
                                                        "" + ";" + //codigocliente
                                                        "CARNE CDC" + ";"   //produto
                                                        );

                                        lsttmp[seq].Add(Parcela);

                                        // decide entre as 3 CAPAS 18,10 ou 08
                                        Parcela = new Parcelas();
                                        Parcela.DadosParcelas.Add(item.Seqcarne.ToString("d8") + ((parcela.Substring(614, 001).ToUpper() == "M"|| parcela.Substring(614, 001).ToUpper() == "S") ? "18" : (Empresa == "01" ? "10" : "08")) 
                                                                                               + item.Qtdeparcela.ToString("d3") + parcela);
                                        lsttmp[seq].Add(Parcela);

                                        // decide entre os 3 MIOLOS 02, 02 ou 19
                                        Parcela = new Parcelas();
                                        Parcela.DadosParcelas.Add(item.Seqcarne.ToString("d8") + ((parcela.Substring(614, 001).ToUpper() == "M"|| parcela.Substring(614, 001).ToUpper() == "S") ? "19" : "02") 
                                                                                               + item.Qtdeparcela.ToString("d3") + parcela);
                                        lsttmp[seq].Add(Parcela);

                                        #region Parcelas Dados
                                        string DadosParcelas = "";

                                        foreach (var dadosparcela in item.Parcelas)
                                        {
                                            DadosParcelas += (dadosparcela.Substring(23, 02).PadRight(14) + (dadosparcela.Substring(39, 02) + "/" + dadosparcela.Substring(37, 02) + "/" + dadosparcela.Substring(33, 04)).PadRight(14) + dadosparcela.Substring(1801, 0014)).PadRight(42);
                                        }

                                        #endregion

                                        //if (parcela.Substring(614, 001).ToUpper() == "M")
                                        //{
                                        //}
                                        //else if (parcela.Substring(614, 001).ToUpper() != "N")
                                        //{
                                        //    Parcela = new Parcelas();
                                        //    Parcela.DadosParcelas.Add((item.Seqcarne.ToString("d8") + "07" + item.Qtdeparcela.ToString("d3") + parcela).PadRight(6562) + DadosParcelas);
                                        //    lsttmp[seq].Add(Parcela);
                                        //}
                                        if (parcela.Substring(614, 001).ToUpper() == "S"  )
                                        {
                                            Parcela = new Parcelas();
                                            Parcela.DadosParcelas.Add((item.Seqcarne.ToString("d8") + "07" + item.Qtdeparcela.ToString("d3") + parcela).PadRight(6562) + DadosParcelas);
                                            lsttmp[seq].Add(Parcela);
                                        }

                                    }

                                    // 03 = Parcelas é = p/todos
                                    Parcela = new Parcelas();
                                    Parcela.DadosParcelas.Add(item.Seqcarne.ToString("d8") + "03" + item.Qtdeparcela.ToString("d3") + parcela);
                                    lsttmp[seq].Add(Parcela);
                                    qtdeparcela++;

                                    // 01 = contra-capa - ja deu a quantidade de laminas?  
                                    if (laminas == item.Parcelas.Count)
                                    {
                                        Parcela = new Parcelas();
                                        Parcela.DadosParcelas.Add(item.Seqcarne.ToString("d8") + "01" + item.Qtdeparcela.ToString("d3") + parcela);
                                        lsttmp[seq].Add(Parcela);
                                    }
                                }
                                //qtdeparcela = item.Qtdeparcela+2;
                                seq++;
                            }

                            #region Ultmimo Carne
                            int prcx = lsttmp[0].Count;
                            if (seq == 1)
                            {
                                if (prcx > 0)
                                {
                                    RelatorioFinal.CarnesBranco++;
                                    RelatorioFinal.LaminasExtra += prcx + 7;
                                }

                                for (int j = 0; j < prcx; j++)
                                {
                                    Parcelas Parcela = new Parcelas();
                                    Parcela.DadosParcelas.Add(lsttmp[0][j].DadosParcelas[0].Substring(00, 20) + "*".PadRight(100, '*'));
                                    lsttmp[1].Add(Parcela);
                                }
                            }

                            for (int j = 0; j < prcx; j++)
                            {
                                for (int i = 0; i < 2; i++)
                                {
                                    sw.WriteLine(lsttmp[i][j].DadosParcelas[0]);
                                }
                            }
                            #endregion

                            //Capa de lote
                            sw.WriteLine("99999999110000000000********************************************************************************" + NM_ARQUIVO.PadRight(30) + SequencialFAC.ToString("d8") + string.Format("{0:dd/MM/yyyy}", DataPostagem).PadRight(25) + DataMovimento.PadRight(15) + Tipocarne.PadRight(20));
                            sw.WriteLine("99999999110000000000********************************************************************************" + NM_ARQUIVO.PadRight(30) + SequencialFAC.ToString("d8"));

                            sw.Dispose();
                            carga.Dispose();
                            Relatorio.Dispose();

                            #region Grava Dados
                            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                            con.Open();

                            string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                            if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "\\Registros.csv"))
                            {
                                Registros(Path.GetDirectoryName(Arquivo) + "\\Schema.ini");
                                OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                                olecom.ExecuteNonQuery();
                                con.Close();

;                               File.Delete(Path.GetDirectoryName(Arquivo) + "\\Registros.csv");
                                File.Delete(Path.GetDirectoryName(Arquivo) + "\\Schema.ini");
                            }

                            #endregion
                        }
                        #endregion

                    }
                }
                File.Delete(Arquivo);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Erro ao processar o arquivo! "+ ncnt.ToString()+" *****  " + ex.StackTrace );
                MessageBox.Show("Erro aqui ... " + ncnt.ToString() + ex.StackTrace  );
            }
        }

        private static bool ArquivoVazio(string Arquivo)
        {
            FileInfo file = new FileInfo(Arquivo);
            if (file.Length == 0)
                return false;
            else
                return true;
        }

        private static bool Verifica_seArquivoEstaVazio(string Arquivo)
        {
            FileInfo file = new FileInfo(Arquivo);
            if (file.Length == 0)
                return true;
            else
                return false;
        }
        private static bool Verifica_seArquivoUtil(string Arquivo)
        {
            FileInfo file = new FileInfo(Arquivo);
            if (file.Length == 0)
                return true;
            else
                return false;
        }


        static void Registros(string Arquivo)
        {
            using (StreamWriter sw = new StreamWriter(Arquivo, false, Encoding.GetEncoding("ISO-8859-1")))
            {
                #region sw Header
                sw.WriteLine("[Registros.csv]");
                sw.WriteLine("ColNameHeader=False");
                sw.WriteLine("Format=Delimited(;)");
                sw.WriteLine("Col1= \"chavepesquisa\" Text Width 255");
                sw.WriteLine("Col2= \"destinatario\" Text Width 255");
                sw.WriteLine("Col3= \"nomekit\" Text Width 255");
                sw.WriteLine("Col4= \"nome\" Text Width 255");
                sw.WriteLine("Col5= \"endereco\" Text Width 255");
                sw.WriteLine("Col6= \"numero\" Text Width 255");
                sw.WriteLine("Col7= \"complemento\" Text Width 255");
                sw.WriteLine("Col8= \"bairro\" Text Width 255");
                sw.WriteLine("Col9= \"cidade\" Text Width 255");
                sw.WriteLine("Col10=\"estado\" Text Width 255");
                sw.WriteLine("Col11=\"cep\" Text Width 255");
                sw.WriteLine("Col12=\"cpfcnpj\" Text Width 255");
                sw.WriteLine("Col13=\"rastreio\" Text Width 255");
                sw.WriteLine("Col14=\"localidade\" Text Width 255");
                sw.WriteLine("Col15=\"tipopostagem\" Text Width 255");
                sw.WriteLine("Col16=\"status\" Text Width 255");
                sw.WriteLine("Col17=\"referencia\" Text Width 255");
                sw.WriteLine("Col18=\"os\" Text Width 255");
                sw.WriteLine("Col19=\"loteid\" Text Width 255");
                sw.WriteLine("Col20=\"audit\" Text Width 255");
                sw.WriteLine("Col21=\"peso\" Text Width 255");
                sw.WriteLine("Col22=\"paginainicial\" Text Width 255");
                sw.WriteLine("Col23=\"paginafinal\" Text Width 255");
                sw.WriteLine("Col24=\"qtdepaginas\" Text Width 255");
                sw.WriteLine("Col25=\"qtdefolhas\" Text Width 255");
                sw.WriteLine("Col26=\"familia\" Text Width 255");
                sw.WriteLine("Col27=\"codinsumo01\" Text Width 255");
                sw.WriteLine("Col28=\"descricao01\" Text Width 255");
                sw.WriteLine("Col29=\"codinsumo02\" Text Width 255");
                sw.WriteLine("Col30=\"descricao02\" Text Width 255");
                sw.WriteLine("Col31=\"codinsumo03\" Text Width 255");
                sw.WriteLine("Col32=\"descricao03\" Text Width 255");
                sw.WriteLine("Col33=\"codinsumo04\" Text Width 255");
                sw.WriteLine("Col34=\"descricao04\" Text Width 255");
                sw.WriteLine("Col35= \"codinsumo05\" Text Width 255");
                sw.WriteLine("Col36= \"descricao05\" Text Width 255");
                sw.WriteLine("Col37= \"crt\" Text Width 255");
                sw.WriteLine("Col38= \"etq\" Text Width 255");
                sw.WriteLine("Col39= \"crn\" Text Width 255");
                sw.WriteLine("Col40= \"manual\" Text Width 255");
                sw.WriteLine("Col41= \"codigomanual\" Text Width 255");
                sw.WriteLine("Col42= \"auditpostagem\" Text Width 255");
                sw.WriteLine("Col43= \"auditkit\" Text Width 255");
                sw.WriteLine("Col44=\"auditado\" Text Width 255");
                sw.WriteLine("Col45=\"fechado\" Text Width 255");
                sw.WriteLine("Col46=\"arquivoentradaid\" Text Width 255");
                sw.WriteLine("Col47=\"datarecepcao\" Text Width 255");
                sw.WriteLine("Col48=\"datapostagem\" Text Width 255");
                sw.WriteLine("Col49=\"dataentrega\" Text Width 255");
                sw.WriteLine("Col50=\"tipokit\" Text Width 255");
                sw.WriteLine("Col51=\"leitura\" Text Width 255");
                sw.WriteLine("Col52=\"email\" Text Width 255");
                sw.WriteLine("Col53=\"contrato\" Text Width 255");
                sw.WriteLine("Col54=\"nomepdf\" Text Width 255");
                sw.WriteLine("Col55=\"telefone\" Text Width 255");
                sw.WriteLine("Col56=\"pathsms\" Text Width 255");
                sw.WriteLine("Col57=\"pathemail\" Text Width 255");
                sw.WriteLine("Col58=\"pdfgerado\" Text Width 255");
                sw.WriteLine("Col59=\"codigocliente\" Text Width 255");
                sw.WriteLine("Col60=\"produto\" Text Width 255");

                sw.Close();
                #endregion
            }
            
        }

        public class Carne
        {
            public string DadosCarne { get; set; }

            public List<String> Parcelas = new List<string>();
            public int Qtdeparcela { get; set; }
            public int Seqcarne { get; set; }
        }
        public class CarneOrdenaCEP
        {
            public string DadosCarne { get; set; }

            public List<String> Parcelas = new List<string>();
            public Int32 CEP { get; set; }
        }
        public class OrdenaCPF
        {
            public string DadosCarne { get; set; }

            public List<String> Parcelas = new List<string>();
            public string CPF { get; set; }
        }


        public class Parcelas
        {
            public List<String> DadosParcelas = new List<string>();
        }

        public class Triagem
        {
            public Int32 cepinicial { get; set; }
            public Int32 cepfinal { get; set; }
            public string cdd { get; set; }
            public string ctc { get; set; }
        }

        public class Relatorio
        {
            public int[] Carnes = new int[100];
            public int CarnesValidos { get; set; }
            public int CarnesBranco { get; set; }
            public int LaminasValidas { get; set; }
            public int LaminasMensagem { get; set; }
            public int LaminasExtra { get; set; }
        }

        private void processarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (Processar(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    // Application.Exit();
                }
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            label3.Visible = true;
            comboBox1.Visible = true;
            comboBox1.Items.Clear();
            button1.Visible = true;
            button2.Visible = true;
            button3.Visible = true;

            #region Pega o Lote
            OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
            connection.Open();
            OleDbDataReader reader = null;
            OleDbCommand command = new OleDbCommand("SELECT * FROM FAC", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add("CARNES " + reader[0].ToString());
            }
            connection.Close();

            connection.Open();
            command = new OleDbCommand("SELECT * FROM FAC_CCB", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add("CCB " + reader[0].ToString());
            }
            connection.Close();

            connection.Open();
            command = new OleDbCommand("SELECT * FROM FAC_CQC", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add("CQC " + reader[0].ToString());
            }
            connection.Close();

            connection.Open();
            command = new OleDbCommand("SELECT * FROM FAC_COBRANCA", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add("COBRANCA " + reader[0].ToString());
            }
            connection.Close();

            connection.Open();
            command = new OleDbCommand("SELECT * FROM FAC_TC", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add("TC " + reader[0].ToString());
            }
            connection.Close();

            connection.Open();
            command = new OleDbCommand("SELECT * FROM FAC_CVV", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add("CVV " + reader[0].ToString());
            }
            connection.Close();

            connection.Open();
            command = new OleDbCommand("SELECT * FROM FAC_DVVDUPLICIDADE", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add("DVVDUPLICIDADE " + reader[0].ToString());
            }
            connection.Close();

            connection.Open();
            command = new OleDbCommand("SELECT * FROM FAC_BACEN", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add("BACEN " + reader[0].ToString());
            }
            connection.Close();

            connection.Open();
            command = new OleDbCommand("SELECT * FROM FAC_CCT", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add("CCT " + reader[0].ToString());
            }
            connection.Close();
            #endregion          
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(comboBox1.Text))
            {
                DataPostagem = this.dateTimePicker1.Value.Date;

                DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    var cartao = cartaopostagem;
                    if (comboBox1.Text.Contains("CCB"))
                        cartao = cartaopostagem02;
                    else if (comboBox1.Text.Contains("BACEN"))
                        cartao = cartaopostagem02;
                    else if (comboBox1.Text.Contains("CCT"))
                        cartao = cartaopostagem02;
                    else if (comboBox1.Text.Contains("TC"))
                        cartao = cartaopostagem02;
                    else if (comboBox1.Text.Contains("DVVDUPLICIDADE"))
                        cartao = cartaopostagem02;
                    else if (comboBox1.Text.Contains("CQC"))
                        cartao = cartaopostagem02;
                    else if (comboBox1.Text.Contains("CVV"))
                        cartao = cartaopostagem02;
                    else if (comboBox1.Text.Contains("CONSIGNADO"))
                        cartao = cartaopostagem02;
                    else if (comboBox1.Text.Contains("VEICULOS"))
                        cartao = cartaopostagem02;
                    else if (comboBox1.Text.Contains("COBRANCA"))
                        cartao = cartaopostagem03;

                    ArquivoMidiaOS(comboBox1.Text.Replace("CARNES", "").Replace("CCT", "").Replace("CCB", "").Trim().Replace("COBRANCA", "").Trim().Replace("TC", "").Trim().Replace("CQC", "").Trim().Replace("CVV", "").Trim().Replace("DVVDUPLICIDADE", "").Trim().Replace("BACEN", "").Trim().Trim().Replace("CONSIGNADO", "").Trim().Trim().Replace("VEICULOS", "").Trim(), cartao);
                    GerarListaOS(comboBox1.Text.Replace("CARNES", "").Replace("CCT", "").Replace("CCB", "").Trim().Replace("COBRANCA", "").Trim().Replace("TC", "").Trim().Replace("CQC", "").Trim().Replace("CVV", "").Trim().Replace("DVVDUPLICIDADE", "").Trim().Replace("BACEN", "").Trim().Trim().Replace("CONSIGNADO", "").Trim().Trim().Replace("VEICULOS", "").Trim(), cartao);
                    MessageBox.Show("Midias processadas com sucesso.");
                    //Application.Exit();
                }
            }
        }

        private static MemoryStream ArquivoMidiaOS(string os, string cartao)
        {
            MemoryStream tmp = new MemoryStream();
            try
            {
                var lote = os;

                string myConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;data source=" + ArquivoMDB;

                OleDbConnection myConnection = new OleDbConnection(myConnectionString);
                myConnection.Open();

                // Execute Queries
                OleDbCommand cmd = myConnection.CreateCommand();
                cmd.CommandText = "SELECT * FROM REGISTROS WHERE OS = '" + os + "' AND STATUS = 'Processado'";
                OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                DataTable myDataTable = new DataTable();
                myDataTable.Load(reader);

                var myEnumerable = myDataTable.AsEnumerable();

                List<BaseResgistros> lista = (from item in myEnumerable
                                              select new BaseResgistros
                                              {
                                                  cif = item.Field<string>("rastreio"),
                                                  peso = item.Field<string>("peso"),
                                                  cep = item.Field<string>("cep")
                                              }).ToList();

                string arquivo = codigoadmistrativo + "_" + lote.PadLeft(5, '0') + "_UNICA_" + drpostagem + ".TXT";

                StreamWriter tw = new StreamWriter(ArquivoMidias + arquivo, false, Encoding.GetEncoding("iso-8859-1"));
                tw.WriteLine("1" + //Tipo da Linha
                             drpostagem + //Codigo Administrativo
                             codigoadmistrativo + //Codigo Administrativo
                             cartao + //Codigo Cartão de Postagem
                             lote + //Numero do Lote de Postagem
                             codigopostagem + //Codigo Postagem
                             ceppostagem + //Cep Postagem
                             contratofac);        //Contrato FAC 

                double pesototal = 0;
                int seqtotal = 0;
                var linha = "";

                foreach (var lst in lista)
                {
                    try
                    {
                        Int64 sequencial = Convert.ToInt64(lst.cif.Substring(15, 11));
                        double peso = Convert.ToDouble(lst.peso);
                        pesototal += peso;

                        string destino = "82031";
                        string cep = lst.cep.Trim().Replace("-", "").Replace(".", "").PadLeft(8, '0');

                        if (lst.cif.Substring(26, 01) == "1")
                            destino = "82015";
                        else if (lst.cif.Substring(26, 01) == "2")
                            destino = "82023";

                        linha = "2" + //Tipo da Linha
                                string.Format("{0:00000000000}", sequencial) + //Sequencial
                                string.Format("{0:N}", peso).Replace(".", "").Replace(",", "").PadLeft(6, '0') + //Peso
                                cep + //Cep
                                destino;

                        if (linha.Length != 31)
                        {
                            tw.Flush();
                            return null;
                        }
                        else
                        {
                            seqtotal++;
                            tw.WriteLine(linha);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(linha);
                    }
                }

                tw.WriteLine("4" + //Tipo da Linha
                             string.Format("{0:0000000}", seqtotal) + //Total de Objetos
                             string.Format("{0:N}", pesototal).Replace(".", "").Replace(",", "").PadLeft(10, '0'));   //Total de Peso
                tw.Flush();

                return tmp;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private static bool GerarListaOS(string os, string cartao)
        {
            var osdasmid = os;
            var lotefac = "";
            var datacif = "";

            OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
            connection.Open();
            OleDbDataReader reader = null;
            OleDbCommand command = new OleDbCommand("SELECT * FROM REGISTROS WHERE OS = '" + os + "' AND STATUS = 'Processado'", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                lotefac = reader["rastreio"].ToString().Substring(10, 05);
                datacif = reader["rastreio"].ToString().Substring(28, 02) + "/" + reader["rastreio"].ToString().Substring(30, 02) + "/20" + reader["rastreio"].ToString().Substring(32, 02);
            }
            connection.Close();


            string myConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;data source=" + ArquivoMDB;

            OleDbConnection myConnection = new OleDbConnection(myConnectionString);
            myConnection.Open();

            // Execute Queries
            OleDbCommand cmd = myConnection.CreateCommand();
            cmd.CommandText = "SELECT * FROM REGISTROS WHERE OS = '" + os + "' AND STATUS = 'Processado'";
            reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

            DataTable myDataTable = new DataTable();
            myDataTable.Load(reader);

            var myEnumerable = myDataTable.AsEnumerable();

            List<BaseResgistros> lista = (from item in myEnumerable
                                          select new BaseResgistros
                                          {
                                              cif = item.Field<string>("rastreio"),
                                              peso = item.Field<string>("peso"),
                                              cep = item.Field<string>("cep"),
                                              lote = item.Field<string>("rastreio").Substring(010, 005),
                                              localidade = item.Field<string>("localidade"),
                                              produto = item.Field<string>("produto"),
                                          }).ToList();

            produto = lista[0].produto;
            var result = lista.GroupBy(o => new { o.lote, o.localidade, o.peso }).Select(grp => new { LOTE = grp.Key, PESO = grp.Sum(o => Convert.ToDouble(o.peso)), TOTAL = grp.Count() });

            var nomearquivopdf = ArquivoMidias + "Lista_" + osdasmid.ToString().PadLeft(6, '0') + "_" + lotefac.ToString().PadLeft(5, '0') + ".PDF";

            Document document = new Document(PageSize.A4, 25, 25, 30, 30);
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(nomearquivopdf, FileMode.Create));
            document.Open();

            string chartLoc = Path.GetDirectoryName(ArquivoMDB) + "\\LISTA_SIMPLES_CORREIO.png";
            iTextSharp.text.Image chartImg = iTextSharp.text.Image.GetInstance(chartLoc);
            chartImg.ScaleAbsolute(610f, 815f);
            chartImg.Alignment = iTextSharp.text.Image.UNDERLYING;
            chartImg.SetAbsolutePosition(-15, 0);
            document.Add(chartImg);

            int QTDETOTAL = 0;
            double PESOTOTAL = 0;

            // the pdf content
            PdfContentByte cb = writer.DirectContent;

            // select the font properties
            BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(BaseColor.BLACK);
            cb.SetFontAndSize(bf, 8);

            cb.BeginText();

            cb.ShowTextAligned(0, lotefac.ToString() + "/" + DataPostagem.Year, 400, 760, 0);
            cb.ShowTextAligned(0, string.Format("{0:dd/MM/yyyy}", DataPostagem), 500, 760, 0);

            cb.ShowTextAligned(0, cliente, 050, 730, 0);

            cb.ShowTextAligned(0, contratofac, 050, 705, 0);
            cb.ShowTextAligned(0, codigoadmistrativo, 270, 705, 0);
            cb.ShowTextAligned(0, cartao, 450, 705, 0);

            cb.ShowTextAligned(0, "TOUCH GRAF SOLUÇÕES GRAFICAS", 050, 680, 0);
            cb.ShowTextAligned(0, "09.469.029/0001-91", 430, 680, 0);

            cb.ShowTextAligned(0, "SPM-SP", 050, 660, 0);
            cb.ShowTextAligned(0, "CTC JAGUARE / GCCAP 3", 210, 660, 0);
            cb.ShowTextAligned(0, codigopostagem, 450, 660, 0);

            #region Obejtos e Pesos Local
            int posiylocal = 458;
            int atexlocal = 0;
            int qtdelocal = 0;
            double pesolocal = 0;

            int posiyestadual = 458;
            int atexestadual = 0;
            int qtdeestadual = 0;
            double pesoestadual = 0;

            int posiynacional = 458;
            int atexnacional = 0;
            int qtdenacional = 0;
            double pesonacional = 0;

            // select the font properties
            bf = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(BaseColor.BLACK);
            cb.SetFontAndSize(bf, 10);
            cb.ShowTextAligned(1, "82015 - Simples Local", 120, 505, 0);
            cb.ShowTextAligned(1, "82023 - Simples Estadual", 295, 505, 0);
            cb.ShowTextAligned(1, "82031 - Simples Nacional", 470, 505, 0);

            // select the font properties
            bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(BaseColor.BLACK);
            cb.SetFontAndSize(bf, 8);

            #region
            foreach (var lst in result)
            {
                if (lst.LOTE.localidade == "1")
                {
                    cb.ShowTextAligned(1, lst.LOTE.peso.ToString(), 052, posiylocal, 0);
                    cb.ShowTextAligned(1, lst.TOTAL.ToString("n0"), 120, posiylocal, 0);
                    cb.ShowTextAligned(1, lst.PESO.ToString("N"), 182, posiylocal, 0);
                    if (atexlocal % 15 == 0 & atexlocal > 0)
                        posiylocal = posiylocal - 18;
                    else
                        posiylocal = posiylocal - 15;
                    atexlocal++;

                    qtdelocal += lst.TOTAL;
                    pesolocal += lst.PESO;
                }
                else if (lst.LOTE.localidade == "2")
                {
                    cb.ShowTextAligned(1, lst.LOTE.peso.ToString(), 238, posiyestadual, 0);
                    cb.ShowTextAligned(1, lst.TOTAL.ToString("n0"), 295, posiyestadual, 0);
                    cb.ShowTextAligned(1, lst.PESO.ToString("N"), 350, posiyestadual, 0);
                    if (atexestadual % 15 == 0 & atexestadual > 0)
                        posiyestadual = posiyestadual - 18;
                    else
                        posiyestadual = posiyestadual - 15;
                    atexestadual++;

                    qtdeestadual += lst.TOTAL;
                    pesoestadual += lst.PESO;
                }
                else if (lst.LOTE.localidade == "3")
                {
                    cb.ShowTextAligned(1, lst.LOTE.peso.ToString(), 415, posiynacional, 0);
                    cb.ShowTextAligned(1, lst.TOTAL.ToString("n0"), 470, posiynacional, 0);
                    cb.ShowTextAligned(1, lst.PESO.ToString("N").ToString(), 530, posiynacional, 0);
                    if (atexnacional % 15 == 0 & atexnacional > 0)
                        posiynacional = posiynacional - 18;
                    else
                        posiynacional = posiynacional - 15;
                    atexnacional++;

                    qtdenacional += lst.TOTAL;
                    pesonacional += lst.PESO;
                }

                QTDETOTAL += lst.TOTAL;
                PESOTOTAL += lst.PESO;
            }
            #endregion

            #endregion

            #region Pesos Totais
            cb.ShowTextAligned(1, qtdelocal.ToString("n0"), 120, 250, 0);
            cb.ShowTextAligned(1, pesolocal.ToString("N"), 182, 250, 0);

            cb.ShowTextAligned(1, qtdeestadual.ToString("n0"), 295, 250, 0);
            cb.ShowTextAligned(1, pesoestadual.ToString("N"), 350, 250, 0);

            cb.ShowTextAligned(1, qtdenacional.ToString("n0"), 470, 250, 0);
            cb.ShowTextAligned(1, pesonacional.ToString("N"), 530, 250, 0);
            #endregion

            #region Numeros de Lote
            int posix = 120;
            int posiy = 190;
            int atex = 0;

            foreach (var lst in result)
            {
                atex++;
                cb.ShowTextAligned(0, lst.LOTE.lote.Trim(), posix, posiy, 0);
                if (atex < 14)
                {
                    posix = posix + 30;
                }
                else
                {
                    atex = 0;
                    posiy = posiy - 9;
                    posix = 120;
                }
                break;
            }
            #endregion

            // select the font properties
            bf = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(BaseColor.BLACK);
            cb.SetFontAndSize(bf, 9);

            cb.ShowTextAligned(0, QTDETOTAL.ToString("n0"), 470, 235, 0);
            cb.ShowTextAligned(0, PESOTOTAL.ToString("N"), 470, 220, 0);

            // select the font properties
            bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            cb.SetColorFill(BaseColor.BLACK);
            cb.SetFontAndSize(bf, 8);

            cb.ShowTextAligned(0, "OS. " + osdasmid + " - " + produto, 100, 160, 0);
            cb.ShowTextAligned(0, "Data CIF: " + datacif, 420, 160, 0);

            cb.ShowTextAligned(0, "Total Postado: " + QTDETOTAL.ToString("n0"), 100, 145, 0);
            cb.ShowTextAligned(0, "Peso Total: " + PESOTOTAL.ToString("N") + " (g)", 420, 145, 0);
            cb.EndText();

            document.Close();
            writer.Close();

            return true;
        }

        public class BaseResgistros
        {
            public string cif { get; set; }
            public string peso { get; set; }
            public string cep { get; set; }
            public string lote { get; set; }
            public string localidade { get; set; }
            public string produto { get; set; }

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            string valor = Interaction.InputBox("Informe o Contrato", "Contrato a Localizar", null, 100, 200);

            if (!string.IsNullOrEmpty(valor))
            {
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("UPDATE REGISTROS SET STATUS = 'RETIDO' WHERE CHAVEPESQUISA = '" + valor + "'", connection);
                reader = command.ExecuteReader();
                connection.Close();
                connection.Open();
                command = new OleDbCommand("SELECT CHAVEPESQUISA FROM REGISTROS WHERE CHAVEPESQUISA = '" + valor + "' AND STATUS = 'RETIDO'", connection);
                reader = command.ExecuteReader();
                bool verifica = false;
                while (reader.Read())
                {
                    verifica = true;
                }

                if (verifica)
                    MessageBox.Show("Contrato " + valor + " localizado e alterado.");
                else
                    MessageBox.Show("Contrato " + valor + " não localizado.");
            }
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            label3.Visible = true;
            comboBox1.Visible = true;
            comboBox1.Items.Clear();
            button1.Visible = true;
            button2.Visible = true;
            button3.Visible = true;

            #region Pega o Lote
            OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
            connection.Open();
            OleDbDataReader reader = null;
            OleDbCommand command = new OleDbCommand("SELECT * FROM FAC", connection);
            reader = command.ExecuteReader();
            while (reader.Read())
            {
                comboBox1.Items.Add(reader[0].ToString());
            }
            connection.Close();
            #endregion

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(comboBox1.Text))
            {
                ArquivoCapa(comboBox1.Text);
                MessageBox.Show("Midias processadas com sucesso.");
                // Application.Exit();
            }
        }

        private static MemoryStream ArquivoCapa(string os)
        {
            MemoryStream tmp = new MemoryStream();
            try
            {

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\TriagemCapa.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), ctc = linha[2] });
                    }
                }
                #endregion

                var lote = os;
                var datafac = "";

                string myConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;data source=" + ArquivoMDB;

                OleDbConnection myConnection = new OleDbConnection(myConnectionString);
                myConnection.Open();

                // Execute Queries
                OleDbCommand cmd = myConnection.CreateCommand();

                cmd.CommandText = "SELECT TOP 1 RASTREIO FROM REGISTROS WHERE OS = '" + os + "' AND STATUS = 'Processado'";
                OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                DataTable myDataTable = new DataTable();
                myDataTable.Load(reader);
                var myEnumerable = myDataTable.AsEnumerable();
                foreach (var item in myEnumerable)
                {
                    datafac = item[0].ToString().Substring(28, 02) + "/" + item[0].ToString().Substring(30, 02) + "/20" + item[0].ToString().Substring(32, 02);
                }
                myConnection.Close();
                myConnection.Open();

                cmd.CommandText = "SELECT DISTINCT(CEP) AS CEP FROM REGISTROS WHERE OS = '" + os + "' AND STATUS = 'Processado'";
                reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                myDataTable = new DataTable();
                myDataTable.Load(reader);

                myEnumerable = myDataTable.AsEnumerable();

                List<Triagem> ListaCapa = new List<Triagem>();
                string arquivo = lote + "_CAPA.TXT";
                StreamWriter tw = new StreamWriter(ArquivoMidias + arquivo, false, Encoding.GetEncoding("iso-8859-1"));

                foreach (var item in myEnumerable)
                {
                    var CEP = Convert.ToInt32(item[0].ToString());
                    var testa = ListaCapa.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();
                    if (testa == null)
                    {
                        testa = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();
                        ListaCapa.Add(new Triagem { ctc = testa.ctc, cepinicial = testa.cepinicial, cepfinal = testa.cepfinal });
                        tw.WriteLine("01|" + datafac +
                                       "|" + testa.cepinicial.ToString().PadLeft(8, '0').Substring(000, 005) + "-" + testa.cepinicial.ToString().PadLeft(8, '0').Substring(005, 003) +
                                       "|" + testa.cepfinal.ToString().PadLeft(8, '0').Substring(000, 005) + "-" + testa.cepfinal.ToString().PadLeft(8, '0').Substring(005, 003) +
                                       "|" + testa.ctc);
                    }
                }

                tw.Flush();

                return tmp;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartas(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    // Application.Exit();
                }
            }
        }

        private static bool ProcessarCartas(string Arquivo)
        {
            try
            {
                var seqarq = 0;
                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\CCB\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCarta(Arquivo, "CCB");
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;

                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CCB", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_CCB (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CCB", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = null;
                StreamWriter carga = new StreamWriter(diretorio + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter relatorio = new StreamWriter(diretorio + "//EPF_RECIBO_GRAFICA.txt", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();
                    SequencialFAC = 0;

                    while (!sr.EndOfStream)
                    {
                        if (SequencialFAC % 1500 == 0)
                        {
                            if (seqarq > 0)
                            {
                                sw.Dispose();
                            }

                            seqarq++;
                            var prearquivo = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(Arquivo)) + "_CCB_" + seqarq.ToString("d3");
                            NM_ARQUIVO = diretorio + Path.ChangeExtension(prearquivo, "SAI");
                            sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                        }

                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[6].Replace("-", "").Trim() == "" ? "95096100" : linhax[6].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        string[] NR = linhax[3].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "CARTAS CCB";

                        sw.WriteLine(linha + ";" +
                                     nmCDD.Trim() + ";" +
                                     CIF.Trim() + ";" +
                                     datamatix);

                        relatorio.WriteLine(linhax[0].Trim() + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DateTime.Now) + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DataPostagem) + ";" +
                                            Lote.ToString("d5") + ";");

                        carga.WriteLine(linhax[0].Trim() + ";" + //chavepesquisa
                                        linhax[2].Trim() + ";" + //destinatario
                                        linhax[2].Trim() + ";" + //nomekit
                                        linhax[2].Trim() + ";" + //nome
                                        linhax[3].Trim() + ";" + //endereco
                                        numeroresidencia + ";" + //numero
                                        linhax[4].Trim() + ";" + //complemento
                                        "" + ";" + //bairro
                                        linhax[5].Trim() + ";" + //cidade
                                        "" + ";" + //estado
                                        linhax[6].Trim() + ";" + //cep
                                        linhax[0].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (18.80).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "6" + ";" + //qtdepaginas
                                        "3" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        "" + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA CCB" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();
                relatorio.Dispose();
                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        private static bool ProcessarCartasCB7(string Arquivo)
        {
            try
            {
                var seqarq = 0;
                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\CB7\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCarta(Arquivo, "CB7");
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;

                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CCB", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_CCB (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CCB", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = null;
                StreamWriter carga = new StreamWriter(diretorio + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter relatorio = new StreamWriter(diretorio + "//EPF_RECIBO_GRAFICA.txt", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();
                    SequencialFAC = 0;

                    while (!sr.EndOfStream)
                    {
                        //if (SequencialFAC % 1500 == 0)
                        if (SequencialFAC % 625 == 0)
                        {
                            if (seqarq > 0)
                            {
                                sw.Dispose();
                            }

                            seqarq++;
                            var prearquivo = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(Arquivo)) + "_CB7_" + seqarq.ToString("d3");
                            NM_ARQUIVO = diretorio + Path.ChangeExtension(prearquivo, "SAI");
                            sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                        }

                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[6].Replace("-", "").Trim() == "" ? "95096100" : linhax[6].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        string[] NR = linhax[3].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "CARTAS CB7";

                        sw.WriteLine(linha + ";" +
                                     nmCDD.Trim() + ";" +
                                     CIF.Trim() + ";" +
                                     datamatix);

                        relatorio.WriteLine(linhax[0].Trim() + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DateTime.Now) + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DataPostagem) + ";" +
                                            Lote.ToString("d5") + ";");

                        
                        carga.WriteLine(linhax[0].Trim() + ";" + //chavepesquisa
                                        linhax[2].Trim() + ";" + //destinatario
                                        linhax[2].Trim() + ";" + //nomekit
                                        linhax[2].Trim() + ";" + //nome
                                        linhax[3].Trim() + ";" + //endereco
                                        numeroresidencia + ";" + //numero
                                        linhax[4].Trim() + ";" + //complemento
                                        "" + ";" + //bairro
                                        linhax[5].Trim() + ";" + //cidade
                                        "" + ";" + //estado
                                        linhax[6].Trim() + ";" + //cep
                                        linhax[0].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (23.60).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "8" + ";" + //qtdepaginas
                                        "4" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        "" + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA CB7" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();
                relatorio.Dispose();
                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }


        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {

        }

        private static bool ProcessarCartasCQC(string Arquivo)
        {
            try
            {
                NM_ARQUIVO = Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "SAI"));

                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartaCQC(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;
                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CQC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_CQC (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CQC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();
                    SequencialFAC = 0;
                    var registro = 0;

                    while (!sr.EndOfStream)
                    {
                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');

                        linhax[03] = linhax[03].Trim();
                        linhax[04] = linhax[04].Trim();
                        linhax[05] = linhax[05].Trim();

                        var tmp = linhax[6].Replace("-", "").Replace(".", "").Trim().PadLeft(8, '0');
                        linhax[6] = tmp.Substring(000, 005) + "-" + tmp.Substring(005, 003);

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[6].Replace("-", "").Trim() == "" ? "95096100" : linhax[6].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        registro++;

                        string[] NR = linhax[3].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "CQC";

                        sw.WriteLine(registro.ToString("d8") + "^" +
                                     string.Join("^", linhax) + "^" +
                                     nmCDD.Trim() + "^" +
                                     CIF.Trim() + "^" +
                                     datamatix + "^"
                                     );

                        carga.WriteLine(linhax[1].Trim() + ";" + //chavepesquisa
                                        linhax[1].Trim() + ";" + //destinatario
                                        "CQC" + ";" + //nomekit
                                        linhax[2].Trim() + ";" + //nome
                                        linhax[3].Trim() + ";" + //endereco
                                        "" + ";" + //numero
                                        "" + ";" + //complemento
                                        "" + ";" + //bairro
                                        linhax[4].Trim() + ";" + //cidade
                                        linhax[5].Trim() + ";" + //estado
                                        linhax[6].Trim() + ";" + //cep
                                        linhax[1].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (4.80).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "5" + ";" + //qtdepaginas
                                        "5" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        "" + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA CQC" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();

                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch
            {
                return false;
            }

        }

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasCQC(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }
        }

        private void descompactarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var sourceFile = openFileDialog1.FileName;

                var destFile = Path.ChangeExtension(sourceFile, ".zip");

                using (var inStream = new FileStream(sourceFile, FileMode.Open))
                {
                    using (var outStream = new MemoryStream())
                    {
                        // tamanho da chave simétrica criptografada
                        var tamanho = new byte[4];
                        inStream.Read(tamanho, 0, 4);

                        var tamChave = BitConverter.ToInt32(tamanho, 0);

                        // lê a chave simétrica criptografada
                        var chaveEnc = new byte[tamChave];
                        inStream.Read(chaveEnc, 0, chaveEnc.Length);

                        // usa o certificado para descriptografar a chave simétrica
                        var chave = DecryptRequest(chaveEnc);
                        
                        // descriptografa utilizando TripleDES com chave aleatória
                        Criptografador.Descriptografar(inStream, outStream, chave, tamChave + 4  );

                        /*
                        // gera o arquivo de saída final
                        using (var decStream = new FileStream(destFile, FileMode.CreateNew))
                        {
                            // grava o conteúdo descriptografado
                            outStream.Seek(0, SeekOrigin.Begin);
                            outStream.CopyTo(decStream);
                        }
                        */
                    }
                }
                MessageBox.Show("Arquivo descriptografado em " + sourceFile);
            }
        }

        private static X509Certificate2 ObterCertificado(string numeroDeSerie)
        {
            if (string.IsNullOrWhiteSpace(numeroDeSerie))
            {
                throw new ArgumentException("O número de série do certificado não pode estar vazio.", "numeroDeSerie");
            }

            var store = new X509Store();
            try
            {
                store.Open(OpenFlags.ReadOnly);

                var certificadosEncontrados = store
                    .Certificates
                    .Cast<X509Certificate2>()
                    .Where(c => String.Equals(c.GetSerialNumberString(), numeroDeSerie, StringComparison.InvariantCultureIgnoreCase))
                    .ToList();

                int exemplo = certificadosEncontrados.Count;

                if (certificadosEncontrados.Count >= 1)
                {
                    string valor = Interaction.InputBox("digite o indice ", "tot. indices: " + exemplo.ToString(), "0", 100, 100);
                    int nIndice = int.Parse(valor);

                    MessageBox.Show("* Tem mais de um certificado " + exemplo.ToString()   ) ;
                    MessageBox.Show(" 1- " + certificadosEncontrados[nIndice].ToString() + "\r\n"  );


                    return certificadosEncontrados[ nIndice ];

                }

                return certificadosEncontrados[0];

                /*
                 * 

                              if (certificadosEncontrados.Count >= 1)
                 {
                     int n = 0;
                     for (n = 0; n < certificadosEncontrados.Count; n++)
                     {
                         if (certificadosEncontrados[n].ToString().Contains(numeroDeSerie))
                         {
                             MessageBox.Show("*** ENCONTRADO: *** \r\n\r\n" + certificadosEncontrados[n].ToString());
                             return certificadosEncontrados[n];
                         }
                     }
                     MessageBox.Show("*** NÃO ENCONTRADO CERTIF. ***");
                     return certificadosEncontrados[0];
                 }
                 else
                 {
                     var mensagem =
                         string.Format("Nenhum certificado válido encontrado nesta máquina. Número de série: {0}",
                             numeroDeSerie);

                     throw new Exception(mensagem);
                 }


                 * */

            }
            finally
            {
                store.Close();
            }
        }



        public static X509Certificate2 GetCertificate()
        {            
            X509Certificate2 cert2 = ObterCertificado("7421d5b03b2ed8f4f76e9c23c7278105");

            return cert2;
        }

        public static byte[] EncryptRequest(byte[] bytes)
        {
            var cert = GetCertificate();
            var provider = cert.PublicKey.Key as RSACryptoServiceProvider;
            var data = provider.Encrypt(bytes, false);

            return data;
        }

        public static byte[] DecryptRequest(byte[] bytes)
        {
            var cert = GetCertificate();
            var provider = cert.PrivateKey as RSACryptoServiceProvider;
            var data = provider.Decrypt(bytes, false);

            return data;
        }

        internal static int ModCep2D(string seq)
        {
            int d, s = 0, p = 1, r;

            for (int i = seq.Length; i > 0; i--)
            {
                r = (Convert.ToInt32(seq.Substring(i - 1, 1)) * p);
                s += r;
            }
            d = ((10 - (s % 10)) % 10);
            return d;
        }

        private void separaErroToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var arquivobanco = openFileDialog1.FileName;
                var arquivocliente = Path.ChangeExtension(openFileDialog1.FileName, ".CSV");
                var arquivonovo = Path.ChangeExtension(openFileDialog1.FileName, "_REMESSA.CSV");

                var objeto = 0;

                StreamWriter sw = new StreamWriter(arquivonovo, false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(arquivobanco, Encoding.GetEncoding("ISO-8859-1")))
                {
                    #region
                    while (!sr.EndOfStream)
                    {
                        var linha = sr.ReadLine();
                        if (linha.Substring(000, 002) == "10")
                        {
                            if (linha.Substring(108, 002) != "02")
                            {
                                var seqremessa = Convert.ToInt32(linha.Substring(063, 008));
                                #region
                                using (StreamReader cliente = new StreamReader(arquivocliente, Encoding.GetEncoding("ISO-8859-1")))
                                {
                                    var linhacliente = cliente.ReadLine();
                                    if (objeto == 0)
                                        sw.WriteLine(linhacliente);
                                    #region
                                    while (!cliente.EndOfStream)
                                    {
                                        linhacliente = cliente.ReadLine();
                                        string[] tmp = linhacliente.Split(';');

                                        if (Convert.ToInt32(tmp[0]) == seqremessa)
                                        {
                                            sw.WriteLine(linhacliente);
                                            break;
                                        }
                                    }
                                    #endregion
                                }
                                #endregion
                                objeto++;
                            }
                        }
                    }
                    #endregion
                }
                sw.Dispose();
                MessageBox.Show("Fim do Processamento!");
            }
        }

        private static bool ProcessarCartasTCCONSIG(string Arquivo)
        {
            try
            {
                NM_ARQUIVO = Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "SAI"));

                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartaTCCONSIG(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;
                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_TC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_TC (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_TC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();

                    SequencialFAC = 0;
                    var registro = 0;

                    while (!sr.EndOfStream)
                    {
                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');
                        var tmp = linhax[10].Replace("-", "").Replace(".", "").Trim().PadLeft(8, '0');
                        linhax[10] = tmp.Substring(000, 005) + "-" + tmp.Substring(005, 003);
                        linhax[02] = Convert.ToUInt64(linhax[02]).ToString(@"000\.000\.000\-00");

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[10].Replace("-", "").Trim() == "" ? "95096100" : linhax[10].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        registro++;

                        string[] NR = linhax[5].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "CONSIG";

                        sw.WriteLine(registro.ToString("d8") + "^" +
                                     string.Join("^", linhax) + "^" +
                                     nmCDD.Trim() + "^" +
                                     CIF.Trim() + "^" +
                                     datamatix + "^");

                        carga.WriteLine(linhax[1].Trim() + ";" + //chavepesquisa
                                        linhax[3].Trim() + ";" + //destinatario
                                        "TCCONSIG" + ";" + //nomekit
                                        linhax[3].Trim() + ";" + //nome
                                        linhax[4].Trim() + ";" + //endereco
                                        linhax[5].Trim() + ";" + //numero
                                        linhax[6].Trim() + ";" + //complemento
                                        linhax[7].Trim() + ";" + //bairro
                                        linhax[8].Trim() + ";" + //cidade
                                        linhax[9].Trim() + ";" + //estado
                                        linhax[10].Trim() + ";" + //cep
                                        linhax[1].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (4.80).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "5" + ";" + //qtdepaginas
                                        "5" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        "" + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA TCCONSIG" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();

                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch
            {
                return false;
            }

        }

        private static void OrdenarCartaTCCONSIG(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();
                sw.WriteLine(linha);

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');
                    Int32 CEP = 0;
                    Int32.TryParse(tmp[10].Replace("-", "").Trim() == "" ? "95096100" : tmp[10].Replace("-", "").Trim(), out CEP);

                    _CarneAtual.CEP = CEP;
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new CarneOrdenaCEP();

                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }

        private void boletoA4ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarBoletoSantaCasa(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }
        }

        private static bool ProcessarBoletoSantaCasa(string Arquivo)
        {
            try
            {
                NM_ARQUIVO = Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "SAI"));

                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarBoletoSantaCasa(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;
                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_COBRANCA", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_COBRANCA (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_COBRANCA", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {

                    SequencialFAC = 0;
                    var registro = 0;

                    while (!sr.EndOfStream)
                    {
                        SequencialFAC++;
                        var linha = sr.ReadLine();

                        var tmp = linha.Substring(326, 008).Replace("-", "").Replace(".", "").Trim().PadLeft(8, '0');
                        tmp = tmp.Substring(000, 005) + "-" + tmp.Substring(005, 003);
                        linha = linha + "^" + tmp.PadRight(10);

                        Int32 CEP = 0;
                        Int32.TryParse(tmp.Replace("-", "").Trim() == "" ? "95096100" : tmp.Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        registro++;

                        string[] NR = linha.Substring(275, 050).Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "BOLETO";

                        sw.WriteLine(linha + "^" +
                                     nmCDD.Trim().PadRight(100).Substring(000, 100) + "^" +
                                     CIF.Trim().PadRight(040).Substring(000, 040) + "^" +
                                     datamatix.PadRight(200).Substring(000, 200) + "^"
                                     );

                        carga.WriteLine(linha.Substring(105, 015).Trim() + ";" + //chavepesquisa
                                        linha.Substring(234, 040).Trim() + ";" + //destinatario
                                        "BOLETO A4" + ";" + //nomekit
                                        linha.Substring(234, 040).Trim() + ";" + //nome
                                        linha.Substring(275, 050).Trim() + ";" + //endereco
                                        "" + ";" + //numero
                                        "" + ";" + //complemento
                                        "" + ";" + //bairro
                                        linha.Substring(334, 015).Trim() + ";" + //cidade
                                        linha.Substring(349, 002).Trim() + ";" + //estado
                                        tmp + ";" + //cep
                                        linha.Substring(220, 014).Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (4.80).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "5" + ";" + //qtdepaginas
                                        "5" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        "" + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "BOLETO SANTA CASA" + ";"   //produto
                                        );
                    }
                }
                sw.Dispose();
                carga.Dispose();

                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void OrdenarBoletoSantaCasa(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    if (linha.Substring(000, 001) == "1")
                    {
                        _CarneAtual.DadosCarne = linha;
                        Int32 CEP = 0;
                        Int32.TryParse(linha.Substring(326, 008).Replace("-", "").Trim() == "" ? "95096100" : linha.Substring(326, 008).Replace("-", "").Trim(), out CEP);

                        _CarneAtual.CEP = CEP;
                        _CarneAtual.Parcelas.Add(linha);
                        _Carnes.Add(_CarneAtual);
                        _CarneAtual = new CarneOrdenaCEP();
                    }
                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasDetran(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }
        }

        private static bool ProcessarCartasDetran(string Arquivo)
        {
            try
            {
                NM_ARQUIVO = Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "SAI"));

                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartasDetran(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;
                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_TC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_TC (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_TC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();

                    SequencialFAC = 0;
                    var registro = 0;

                    while (!sr.EndOfStream)
                    {
                        SequencialFAC++;
                        linha = sr.ReadLine();

                        try
                        {
                            var linhax = linha.Split(';');
                            var tmp = linhax[10].Replace("-", "").Replace(".", "").Trim().PadLeft(8, '0');
                            linhax[10] = tmp.Substring(000, 005) + "-" + tmp.Substring(005, 003);

                            try
                            {
                                if (linhax[03].Trim().Length < 12)
                                    linhax[03] = Convert.ToUInt64(linhax[03]).ToString(@"000\.000\.000\-00");
                                else
                                    linhax[03] = Convert.ToUInt64(linhax[03]).ToString(@"00\.000\.000\/0000\-00");
                            }
                            catch
                            {
                                linhax[03] = linhax[03];
                            }

                            Int32 CEP = 0;
                            Int32.TryParse(linhax[10].Replace("-", "").Trim() == "" ? "95096100" : linhax[10].Replace("-", "").Trim(), out CEP);

                            string Destino = "3";
                            if (CEP < 10000000)
                                Destino = "1";
                            else if (CEP < 20000000)
                                Destino = "2";

                            var nmCDD = "Não Localizado";
                            var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                            if (CDD != null)
                                nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                            var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                            registro++;

                            string[] NR = linhax[5].Trim().Split(' ');
                            string numeroresidencia = "";
                            int nr = NR.Count();

                            try
                            {
                                for (int m = 0; m < nr; m++)
                                {
                                    numeroresidencia = NR[m];
                                    int n = 0;
                                    var isNumeric = int.TryParse(numeroresidencia, out n);
                                    if (isNumeric)
                                    {
                                        numeroresidencia = n.ToString();
                                        break;
                                    }
                                    numeroresidencia = "0";
                                }
                            }
                            catch { }

                            int nx = 0;
                            var isNumericx = int.TryParse(numeroresidencia, out nx);
                            if (!isNumericx)
                                numeroresidencia = "0";

                            string cddestino = "82031";
                            if (Destino == "1")
                                cddestino = "82015";
                            else if (Destino == "2")
                                cddestino = "82023";

                            var datamatix = CEP.ToString("d8") +
                                            numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                            "01310930" +
                                            "02150" +
                                            ModCep2D(CEP.ToString("d8")) +
                                            "01" +
                                            CIF +
                                            "0000000000" +
                                            cddestino +
                                            "000000000000000" +
                                            "006422100" +
                                            "|" +
                                            "DETRAN";

                            sw.WriteLine(registro.ToString("d8") + "^" +
                                         string.Join("^", linhax) + "^" +
                                         nmCDD.Trim() + "^" +
                                         CIF.Trim() + "^" +
                                         datamatix + "^" +
                                         "São Paulo, " + string.Format("{0:dd 'de' MMMM 'de' yyyy}", DateTime.Now) + "^");

                            carga.WriteLine(linhax[1].Trim() + ";" + //chavepesquisa
                                            linhax[2].Trim() + ";" + //destinatario
                                            "DETRAN" + ";" + //nomekit
                                            linhax[2].Trim() + ";" + //nome
                                            linhax[4].Trim() + ";" + //endereco
                                            linhax[5].Trim() + ";" + //numero
                                            linhax[6].Trim() + ";" + //complemento
                                            linhax[7].Trim() + ";" + //bairro
                                            linhax[8].Trim() + ";" + //cidade
                                            linhax[9].Trim() + ";" + //estado
                                            linhax[10].Trim() + ";" + //cep
                                            linhax[3].Trim() + ";" + //cpfcnpj
                                            CIF + ";" + //rastreio
                                            CIF.Substring(26, 01) + ";" + //localidade
                                            "FAC" + ";" + //tipopostagem
                                            "Processado" + ";" + //status
                                            "" + ";" + //referencia
                                            Lote.ToString() + ";" + //os
                                            SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                            SequencialFAC.ToString("d8") + ";" + //audit
                                            (4.80).ToString() + ";" + //peso
                                            "1" + ";" + //paginainicial
                                            "2" + ";" + //paginafinal
                                            "5" + ";" + //qtdepaginas
                                            "5" + ";" + //qtdefolhas
                                            "P" + ";" + //familia
                                            "" + ";" + //codinsumo01
                                            "" + ";" + //descricao01
                                            "" + ";" + //codinsumo02
                                            "" + ";" + //descricao02
                                            "" + ";" + //codinsumo03
                                            "" + ";" + //descricao03
                                            "" + ";" + //codinsumo04
                                            "" + ";" + //descricao04
                                            "" + ";" + //codinsumo05
                                            "" + ";" + //descricao05
                                            "" + ";" + //crt
                                            "" + ";" + //etq
                                            "" + ";" + //crn
                                            "" + ";" + //manual
                                            "" + ";" + //codigomanual
                                            "" + ";" + //auditpostagem
                                            "" + ";" + //auditkit
                                            "" + ";" + //auditado
                                            "" + ";" + //fechado
                                            "" + ";" + //arquivoentradaid
                                            "" + ";" + //datarecepcao
                                            "" + ";" + //datapostagem
                                            "" + ";" + //dataentrega
                                            "" + ";" + //tipokit
                                            "" + ";" + //leitura
                                            "" + ";" + //email
                                            "" + ";" + //contrato
                                            "" + ";" + //nomepdf
                                            "" + ";" + //telefone
                                            "" + ";" + //pathsms
                                            "" + ";" + //pathemail
                                            "" + ";" + //pdfgerado
                                            "" + ";" + //codigocliente
                                            "CARTA DETRAN" + ";"   //produto
                                            );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(linha);
                        }
                    }
                }
                sw.Dispose();
                carga.Dispose();

                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch
            {
                return false;
            }

        }

        private static void OrdenarCartasDetran(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();
                sw.WriteLine(linha);

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');
                    Int32 CEP = 0;
                    Int32.TryParse(tmp[10].Replace("-", "").Trim() == "" ? "95096100" : tmp[10].Replace("-", "").Trim(), out CEP);

                    _CarneAtual.CEP = CEP;
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new CarneOrdenaCEP();

                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasTC(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }
        }

        private static bool ProcessarCartasTC(string Arquivo)
        {
            try
            {
                NM_ARQUIVO = Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "SAI"));

                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartasTC(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;
                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_TC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_TC (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_TC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();

                    SequencialFAC = 0;
                    var registro = 0;

                    while (!sr.EndOfStream)
                    {
                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');
                        var tmp = linhax[11].Replace("-", "").Replace(".", "").Trim().PadLeft(8, '0');
                        linhax[11] = tmp.Substring(000, 005) + "-" + tmp.Substring(005, 003);

                        if (linhax[03].Trim().Length < 12)
                            linhax[03] = Convert.ToUInt64(linhax[03]).ToString(@"000\.000\.000\-00");
                        else
                            linhax[03] = Convert.ToUInt64(linhax[03]).ToString(@"00\.000\.000\/0000\-00");

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[11].Replace("-", "").Trim() == "" ? "95096100" : linhax[11].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        registro++;

                        string[] NR = linhax[6].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "TC";

                        sw.WriteLine(registro.ToString("d8") + "^" +
                                     string.Join("^", linhax) + "^" +
                                     nmCDD.Trim() + "^" +
                                     CIF.Trim() + "^" +
                                     datamatix + "^" +
                                     "São Paulo, " + string.Format("{0:dd 'de' MMMM 'de' yyyy}", DateTime.Now) + "^");

                        carga.WriteLine(linhax[2].Trim() + ";" + //chavepesquisa
                                        linhax[4].Trim() + ";" + //destinatario
                                        linhax[0].Trim() + ";" + //nomekit
                                        linhax[4].Trim() + ";" + //nome
                                        linhax[5].Trim() + ";" + //endereco
                                        linhax[6].Trim() + ";" + //numero
                                        linhax[7].Trim() + ";" + //complemento
                                        linhax[8].Trim() + ";" + //bairro
                                        linhax[9].Trim() + ";" + //cidade
                                        linhax[10].Trim() + ";" + //estado
                                        linhax[11].Trim() + ";" + //cep
                                        linhax[3].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (9.60).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "5" + ";" + //qtdepaginas
                                        "5" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        linhax[0].Trim() + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA TC" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();

                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch
            {
                return false;
            }

        }

        private static void OrdenarCartasTC(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();
                sw.WriteLine(linha);

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');
                    Int32 CEP = 0;
                    Int32.TryParse(tmp[11].Replace("-", "").Trim() == "" ? "95096100" : tmp[11].Replace("-", "").Trim(), out CEP);

                    _CarneAtual.CEP = CEP;
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new CarneOrdenaCEP();

                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }

        private static void OrdenarCartasDUT(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();
                sw.WriteLine(linha);

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');
                    Int32 CEP = 0;
                    Int32.TryParse(tmp[12].Replace("-", "").Trim() == "" ? "95096100" : tmp[12].Replace("-", "").Trim(), out CEP);

                    _CarneAtual.CEP = CEP;
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new CarneOrdenaCEP();

                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }

        private static void OrdenarCartasBNDU(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();
                sw.WriteLine(linha);

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');
                    Int32 CEP = 0;
                    Int32.TryParse(tmp[11].Replace("-", "").Trim() == "" ? "95096100" : tmp[11].Replace("-", "").Trim(), out CEP);

                    _CarneAtual.CEP = CEP;
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new CarneOrdenaCEP();

                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(comboBox1.Text))
            {
                DataPostagem = this.dateTimePicker1.Value.Date;
                ProcessarRelatorio(comboBox1.Text);
                MessageBox.Show("Relatório processado com sucesso.");
                // Application.Exit();
            }
        }

        private static MemoryStream ProcessarRelatorio(string os)
        {
            MemoryStream tmp = new MemoryStream();
            try
            {
                var lote = os.Split(' ')[1];

                string myConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;data source=" + ArquivoMDB;

                OleDbConnection myConnection = new OleDbConnection(myConnectionString);
                myConnection.Open();

                // Execute Queries
                OleDbCommand cmd = myConnection.CreateCommand();

                cmd.CommandText = "SELECT CHAVEPESQUISA, ENDERECO, NUMERO, COMPLEMENTO, BAIRRO, CIDADE, CEP, PESO, NOMEKIT, NOME, ESTADO, CPFCNPJ, RASTREIO FROM REGISTROS WHERE OS = '" + lote + "' AND STATUS = 'Processado'";
                OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.CloseConnection);

                DataTable myDataTable = new DataTable();
                myDataTable.Load(reader);
                var myEnumerable = myDataTable.AsEnumerable();
                StreamWriter tw = new StreamWriter(ArquivoMidias + os.Replace(' ', '_') + ".CSV", false, Encoding.GetEncoding("iso-8859-1"));

                int registro = 0;
                foreach (var item in myEnumerable)
                {
                    registro++;

                    string datafac = item[12].ToString().Substring(28, 02) + "/" + item[12].ToString().Substring(30, 02) + "/20" + item[12].ToString().Substring(32, 02);

                    tw.WriteLine(registro + ";" +
                                 datafac + ";" +
                                 item[00].ToString().Trim() + ";" +
                                 item[01].ToString().Trim() + ";" +
                                 item[02].ToString().Trim() + ";" +
                                 item[03].ToString().Trim() + ";" +
                                 item[04].ToString().Trim() + ";" +
                                 item[05].ToString().Trim() + ";" +
                                 item[06].ToString().Trim() + ";" +
                                 item[07].ToString().Trim() + ";" +
                                 item[08].ToString().Trim() + ";" +
                                 item[09].ToString().Trim() + ";" +
                                 item[10].ToString().Trim() + ";" +
                                 item[11].ToString().Trim() + ";" +
                                 item[12].ToString().Trim() + ";" +
                                 string.Format("{0:dd/MM/yyyy}", DateTime.Now) + ";" +
                                 string.Format("{0:dd/MM/yyyy}", DataPostagem));
                }
                myConnection.Close();
                myConnection.Open();

                tw.Flush();

                return tmp;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasTCCONSIG(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }
        }

        private void fase01ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                if (ProcessarCartasCVV(openFileDialog1.FileName))
                    MessageBox.Show("Arquivo Processado com sucesso.");
                else
                    MessageBox.Show("Erro ao processar o arquivo");

                Application.Exit();
            }
        }

        private static bool ProcessarCartasCVV(string Arquivo)
        {
            try
            {
                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartaCVVCPF(Arquivo);
                OrdenarCartaCVVCPFOrdena(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileNameWithoutExtension(Arquivo) + ".SA1");
                Arquivo = diretorio + Path.GetFileName(Arquivo);
                StreamWriter CorreioReter = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "RETIDO_CEP")), false, Encoding.GetEncoding("ISO-8859-1"));

                StreamWriter sw = null;
                var registro = 0;
                var seqarq = 0;
                var CEPOK = false;

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    while (!sr.EndOfStream)
                    {
                        var linha = sr.ReadLine();

                        if (linha.Substring(000, 002) == "01")
                        {
                            #region
                            var tmp = linha.Substring(587, 010).Replace("-", "").Replace(".", "").Trim();

                            if (tmp.Length != 8)
                                tmp = "0";

                            Int32 CEP = 0;
                            Int32.TryParse(tmp.Replace("-", "").Trim() == "" ? "00000000" : tmp.Replace("-", "").Trim(), out CEP);

                            CEPOK = false;
                            var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                            if (CDD != null)
                                CEPOK = true;

                            #endregion

                            if (registro % 50000 == 0 && CEPOK)
                            {
                                if (seqarq > 0)
                                    sw.Dispose();
                                seqarq++;
                                NM_ARQUIVO = Path.GetDirectoryName(Arquivo) + "\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "SAI_" + seqarq.ToString("d3")));
                                sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                            }

                            if (CDD != null)
                                registro++;
                        }

                        if (linha.Substring(000, 002) == "01" || linha.Substring(000, 002) == "00")
                        {
                            if (CEPOK)
                                sw.WriteLine(linha);
                            else
                                CorreioReter.WriteLine(linha);
                        }
                    }
                }
                CorreioReter.Dispose();
                sw.Dispose();
                File.Delete(Arquivo);

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool ProcessarCartasCVVFAC(string Arquivo)
        {
            try
            {
                var nrlote = 0;
                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CVV", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_CVV (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CVV", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = new StreamWriter(Arquivo + "_FAC", false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                var cpf = "";

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = "";

                    SequencialFAC = 0;
                    var registro = 0;

                    while (!sr.EndOfStream)
                    {
                        linha = sr.ReadLine();
                        if (linha.Substring(00, 02) == "01")
                        {
                            SequencialFAC++;
                            cpf = linha.Substring(02, 20).Trim();

                            #region
                            var tmp = linha.Substring(587, 010).Replace("-", "").Replace(".", "").Trim().PadLeft(8, '0');
                            linha += tmp.Substring(000, 005) + "-" + tmp.Substring(005, 003);

                            Int32 CEP = 0;
                            Int32.TryParse(tmp.Replace("-", "").Trim() == "" ? "00000000" : tmp.Replace("-", "").Trim(), out CEP);

                            string Destino = "3";
                            if (CEP < 10000000)
                                Destino = "1";
                            else if (CEP < 20000000)
                                Destino = "2";

                            var nmCDD = "Não Localizado";
                            var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                            if (CDD != null)
                                nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                            var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                            registro++;

                            string[] NR = linha.Substring(312, 010).Trim().Split(' ');
                            string numeroresidencia = "";
                            int nr = NR.Count();

                            try
                            {
                                for (int m = 0; m < nr; m++)
                                {
                                    numeroresidencia = NR[m];
                                    int n = 0;
                                    var isNumeric = int.TryParse(numeroresidencia, out n);
                                    if (isNumeric)
                                    {
                                        numeroresidencia = n.ToString();
                                        break;
                                    }
                                    numeroresidencia = "0";
                                }
                            }
                            catch { }

                            int nx = 0;
                            var isNumericx = int.TryParse(numeroresidencia, out nx);
                            if (!isNumericx)
                                numeroresidencia = "0";

                            string cddestino = "82031";
                            if (Destino == "1")
                                cddestino = "82015";
                            else if (Destino == "2")
                                cddestino = "82023";

                            var datamatix = CEP.ToString("d8") +
                                            numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                            "01310930" +
                                            "02150" +
                                            ModCep2D(CEP.ToString("d8")) +
                                            "01" +
                                            CIF +
                                            "0000000000" +
                                            cddestino +
                                            "000000000000000" +
                                            "006422100" +
                                            "|" +
                                            "CVV";

                            sw.WriteLine(registro.ToString("d8") + "^" +
                                         linha + "^" +
                                         nmCDD.Trim().PadRight(150) + "^" +
                                         CIF.Trim() + "^" +
                                         datamatix + "^");

                            carga.WriteLine(linha.Substring(002, 020).Trim() + ";" + //chavepesquisa
                                            linha.Substring(022, 100).Trim() + ";" + //destinatario
                                            "CVV" + ";" + //nomekit
                                            linha.Substring(022, 100).Trim() + ";" + //nome
                                            linha.Substring(212, 100).Trim() + ";" + //endereco
                                            linha.Substring(312, 020).Trim() + ";" + //numero
                                            linha.Substring(332, 050).Trim() + ";" + //complemento
                                            linha.Substring(382, 100).Trim() + ";" + //bairro
                                            linha.Substring(482, 100).Trim() + ";" + //cidade
                                            linha.Substring(582, 005).Trim() + ";" + //estado
                                            linha.Substring(587, 010).Trim() + ";" + //cep
                                            linha.Substring(002, 020).Trim() + ";" + //cpfcnpj
                                            CIF + ";" + //rastreio
                                            CIF.Substring(26, 01) + ";" + //localidade
                                            "FAC" + ";" + //tipopostagem
                                            "Processado" + ";" + //status
                                            "" + ";" + //referencia
                                            Lote.ToString() + ";" + //os
                                            SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                            SequencialFAC.ToString("d8") + ";" + //audit
                                            (4.80).ToString() + ";" + //peso
                                            "1" + ";" + //paginainicial
                                            "2" + ";" + //paginafinal
                                            "5" + ";" + //qtdepaginas
                                            "5" + ";" + //qtdefolhas
                                            "P" + ";" + //familia
                                            "" + ";" + //codinsumo01
                                            "" + ";" + //descricao01
                                            "" + ";" + //codinsumo02
                                            "" + ";" + //descricao02
                                            "" + ";" + //codinsumo03
                                            "" + ";" + //descricao03
                                            "" + ";" + //codinsumo04
                                            "" + ";" + //descricao04
                                            "" + ";" + //codinsumo05
                                            "" + ";" + //descricao05
                                            "" + ";" + //crt
                                            "" + ";" + //etq
                                            "" + ";" + //crn
                                            "" + ";" + //manual
                                            "" + ";" + //codigomanual
                                            "" + ";" + //auditpostagem
                                            "" + ";" + //auditkit
                                            "" + ";" + //auditado
                                            "" + ";" + //fechado
                                            "" + ";" + //arquivoentradaid
                                            "" + ";" + //datarecepcao
                                            "" + ";" + //datapostagem
                                            "" + ";" + //dataentrega
                                            "" + ";" + //tipokit
                                            "" + ";" + //leitura
                                            "" + ";" + //email
                                            "" + ";" + //contrato
                                            "" + ";" + //nomepdf
                                            "" + ";" + //telefone
                                            "" + ";" + //pathsms
                                            "" + ";" + //pathemail
                                            "" + ";" + //pdfgerado
                                            "" + ";" + //codigocliente
                                            "CARTA CVV" + ";"   //produto
                                            );
                            #endregion
                        }
                        else if (linha.Substring(00, 02) == "00")
                        {
                            if (cpf != linha.Substring(02, 20).Trim())
                            {
                                MessageBox.Show("Erro de CPF");
                                MessageBox.Show("Erro de CPF");
                                MessageBox.Show("Erro de CPF");
                                break;
                            }

                            sw.WriteLine(registro.ToString("d8") + "^" +
                                         linha + "^");
                        }

                    }
                }
                sw.Dispose();
                carga.Dispose();

                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #region Valida Arquivo
                bool valida = false;
                var contador = 0;
                var seqarquivo = 0;
                cpf = "";
                StreamWriter saida = null;

                using (StreamReader srx = new StreamReader(Arquivo + "_FAC", Encoding.GetEncoding("ISO-8859-1")))
                {
                    try
                    {
                        while (!srx.EndOfStream)
                        {
                            var linhax = srx.ReadLine();

                            if (linhax.Substring(09, 02) == "01")
                            {
                                cpf = linhax.Substring(11, 20).Trim();
                                if (contador % 10000 == 0)
                                {
                                    if (seqarquivo > 0)
                                        saida.Dispose();
                                    seqarquivo++;
                                    saida = new StreamWriter(Arquivo + "_FAC_" + seqarquivo.ToString("d3"), false, Encoding.GetEncoding("ISO-8859-1"));
                                }
                                contador++;
                            }
                            else if (linhax.Substring(09, 02) == "00")
                            {
                                if (cpf != linhax.Substring(11, 20).Trim())
                                {
                                    MessageBox.Show("Erro de CPF");
                                    MessageBox.Show("Erro de CPF");
                                    MessageBox.Show("Erro de CPF");
                                    valida = false;
                                    break;
                                }
                            }

                            saida.WriteLine(linhax);


                            if (linhax.Length > 100)
                            {
                                if (linhax.Substring(615, 001) != "^" || linhax.Substring(904, 005) != "|CVV^" || linhax.Length != 909)
                                {
                                    valida = false;
                                    break;
                                }
                                else
                                    valida = true;
                            }
                        }
                    }
                    catch { valida = false; }
                }

                saida.Dispose();

                if (valida == false)
                {
                    MessageBox.Show("Erro ao validar o arquivo.");
                    File.Delete(Arquivo + "_FAC");
                }
                else
                {
                    File.Delete(Arquivo + "_FAC");
                }
                #endregion

                #endregion
                return true;
            }
            catch
            {
                return false;
            }

        }

        private static void OrdenarCartaCVVCPF(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<OrdenaCPF> _Carnes = new List<OrdenaCPF>();
                OrdenaCPF _CarneAtual = new OrdenaCPF();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileNameWithoutExtension(Arquivo) + ".SA1", false, Encoding.GetEncoding("ISO-8859-1"));
                sw.WriteLine("33");
                string linha = sr.ReadLine();

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');

                    _CarneAtual.CPF = tmp[0];
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new OrdenaCPF();
                }

                var lista = _Carnes.OrderBy(x => x.CPF);
                var cpfold = "";
                foreach (var item in lista)
                {
                    if (cpfold != item.CPF)
                    {
                        string[] tmp = item.Parcelas[0].Split(';');
                        sw.WriteLine("01" +
                                     tmp[00].Trim().PadRight(020) + //CPF
                                     tmp[01].Trim().PadRight(150) + //NmCliente
                                     tmp[02].Trim().PadRight(020) + //IdContrato
                                     tmp[03].Trim().PadRight(020) + //IdCvv
                                     tmp[04].Trim().PadRight(100) + //Logradouro
                                     tmp[05].Trim().PadRight(020) + //Numero
                                     tmp[06].Trim().PadRight(050) + //Complemento
                                     tmp[07].Trim().PadRight(100) + //Bairro
                                     tmp[08].Trim().PadRight(100) + //Cidade
                                     tmp[09].Trim().PadRight(005) + //Uf
                                     tmp[10].Trim().PadRight(010)); //Cep

                        cpfold = item.CPF;
                    }
                    foreach (var parcela in item.Parcelas)
                    {
                        string[] tmp = parcela.Split(';');
                        sw.WriteLine("00" +
                                     tmp[00].Trim().PadRight(020) + //CPF
                                     tmp[02].Trim().PadRight(020) + //IdContrato
                                     tmp[03].Trim().PadRight(020)); //IdCvv

                        if (cpfold != tmp[00].Trim())
                        {
                            MessageBox.Show("Erro ao Processar.");
                            MessageBox.Show("Erro ao Processar.");
                            MessageBox.Show("Erro ao Processar.");
                        }
                    }
                }
                sw.WriteLine("44");
                sw.Dispose();
            }
        }

        private static void OrdenarCartaCVVCPFOrdena(string Arquivo)
        {
            var contador01 = 0;
            var contador02 = 0;
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();
                string CarneOld = "";

                StreamWriter sw = new StreamWriter(Path.ChangeExtension(Arquivo, "TXT"), false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter relatorio = new StreamWriter(Path.ChangeExtension(Arquivo, "VER"), false, Encoding.GetEncoding("ISO-8859-1"));

                while (!sr.EndOfStream)
                {
                    string linha = sr.ReadLine();

                    try
                    {
                        #region
                        if (linha.Substring(000, 002) == "33")
                        {
                        }
                        #region Tipo 2
                        else if (linha.Substring(000, 002) == "01")
                        {
                            contador01++;
                            if (CarneOld != linha.Substring(003, 020).Trim())
                            {
                                if (!string.IsNullOrEmpty(CarneOld))
                                {
                                    _Carnes.Add(_CarneAtual);
                                    _CarneAtual = new CarneOrdenaCEP();
                                }
                                CarneOld = linha.Substring(003, 020).Trim();
                                _CarneAtual.DadosCarne = CarneOld;
                                _CarneAtual.CEP = Convert.ToInt32(linha.Substring(587, 008).Trim().PadLeft(8, '0'));
                            }
                            _CarneAtual.Parcelas.Add(linha);
                        }
                        else if (linha.Substring(000, 002) == "00")
                        {
                            _CarneAtual.Parcelas.Add(linha);
                        }
                        #endregion

                        #region Tipo 4
                        else if (linha.Substring(000, 002) == "44")
                        {
                            #region Ultimo Carne
                            _Carnes.Add(_CarneAtual);
                            #endregion

                            var lista = _Carnes.OrderByDescending(x => x.CEP);
                            foreach (var item in lista)
                            {
                                var qtde = 0;
                                foreach (var parcela in item.Parcelas)
                                {
                                    if (parcela.Substring(00, 02) == "00")
                                    {
                                        qtde++;
                                    }

                                    if (qtde == 10)
                                    {
                                        sw.WriteLine(item.Parcelas[0]);
                                        qtde = 0;
                                        relatorio.WriteLine(item.Parcelas[0]);
                                        contador02++;
                                    }
                                    sw.WriteLine(parcela);
                                }
                            }

                            //Trailler
                            sw.WriteLine(linha);
                            sw.Dispose();
                            relatorio.WriteLine("Total CPF: " + contador01);
                            relatorio.WriteLine("Total Objetos+9: " + contador02);
                            relatorio.Dispose();
                        }
                        #endregion
                        #endregion
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(linha);
                    }
                }
            }
            File.Delete(Arquivo);
        }

        private void toolStripMenuItem12_Click_1(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasDUT(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }
        }

        private static bool ProcessarCartasDUT(string Arquivo)
        {
            try
            {
                NM_ARQUIVO = Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "SAI"));

                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartasDUT(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;
                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_TC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_TC (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_TC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();

                    SequencialFAC = 0;
                    var registro = 0;

                    while (!sr.EndOfStream)
                    {
                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');
                        var tmp = linhax[12].Replace("-", "").Replace(".", "").Trim().PadLeft(8, '0');
                        linhax[12] = tmp.Substring(000, 005) + "-" + tmp.Substring(005, 003);

                        if (linhax[02].Trim().Length < 12)
                            linhax[02] = Convert.ToUInt64(linhax[02]).ToString(@"000\.000\.000\-00");
                        else
                            linhax[02] = Convert.ToUInt64(linhax[02]).ToString(@"00\.000\.000\/0000\-00");

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[12].Replace("-", "").Trim() == "" ? "95096100" : linhax[12].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        registro++;

                        string[] NR = linhax[7].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "DUT";

                        sw.WriteLine(registro.ToString("d8") + "^" +
                                     string.Join("^", linhax) + "^" +
                                     nmCDD.Trim() + "^" +
                                     CIF.Trim() + "^" +
                                     datamatix + "^" +
                                     "São Paulo, " + string.Format("{0:dd 'de' MMMM 'de' yyyy}", DateTime.Now) + "^");

                        carga.WriteLine(linhax[2].Trim() + ";" + //chavepesquisa
                                        linhax[4].Trim() + ";" + //destinatario
                                        linhax[0].Trim() + ";" + //nomekit
                                        linhax[3].Trim() + ";" + //nome
                                        linhax[6].Trim() + ";" + //endereco
                                        linhax[7].Trim() + ";" + //numero
                                        linhax[8].Trim() + ";" + //complemento
                                        linhax[9].Trim() + ";" + //bairro
                                        linhax[10].Trim() + ";" + //cidade
                                        linhax[11].Trim() + ";" + //estado
                                        linhax[12].Trim() + ";" + //cep
                                        linhax[3].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (4.80).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "5" + ";" + //qtdepaginas
                                        "5" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        linhax[0].Trim() + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA DUT" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();

                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch
            {
                return false;
            }

        }

        private void fase02ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasCVVFAC(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }

        }

        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasBNDU(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }
        }

        private static bool ProcessarCartasBNDU(string Arquivo)
        {
            try
            {
                NM_ARQUIVO = Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "SAI"));

                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartasBNDU(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;
                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_TC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_TC (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_TC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();

                    SequencialFAC = 0;
                    var registro = 0;

                    while (!sr.EndOfStream)
                    {
                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');
                        var tmp = linhax[11].Replace("-", "").Replace(".", "").Trim().PadLeft(8, '0');
                        linhax[11] = tmp.Substring(000, 005) + "-" + tmp.Substring(005, 003);

                        if (linhax[03].Trim().Length < 12)
                            linhax[03] = Convert.ToUInt64(linhax[02]).ToString(@"000\.000\.000\-00");
                        else
                            linhax[03] = Convert.ToUInt64(linhax[02]).ToString(@"00\.000\.000\/0000\-00");

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[11].Replace("-", "").Trim() == "" ? "95096100" : linhax[11].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        registro++;

                        string[] NR = linhax[6].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "BNDU";

                        sw.WriteLine(registro.ToString("d8") + "^" +
                                     string.Join("^", linhax) + "^" +
                                     nmCDD.Trim() + "^" +
                                     CIF.Trim() + "^" +
                                     datamatix + "^" +
                                     "São Paulo, " + string.Format("{0:dd 'de' MMMM 'de' yyyy}", DateTime.Now) + "^");

                        carga.WriteLine(linhax[2].Trim() + ";" + //chavepesquisa
                                        linhax[4].Trim() + ";" + //destinatario
                                        linhax[3].Trim() + ";" + //nomekit
                                        linhax[4].Trim() + ";" + //nome
                                        linhax[5].Trim() + ";" + //endereco
                                        linhax[6].Trim() + ";" + //numero
                                        linhax[7].Trim() + ";" + //complemento
                                        linhax[8].Trim() + ";" + //bairro
                                        linhax[9].Trim() + ";" + //cidade
                                        linhax[10].Trim() + ";" + //estado
                                        linhax[11].Trim() + ";" + //cep
                                        linhax[3].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (4.80).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "5" + ";" + //qtdepaginas
                                        "5" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        linhax[2].Trim() + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA BNDU" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();

                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch
            {
                return false;
            }

        }

        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasDVVDUPLICIDADE(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }
        }
        private static bool ProcessarCartasDVVDUPLICIDADE(string Arquivo)
        {
            try
            {
                NM_ARQUIVO = Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "SAI"));

                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartasDVVDUPLICIDADE(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;
                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_DVVDUPLICIDADE", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_DVVDUPLICIDADE (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_DVVDUPLICIDADE", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();

                    SequencialFAC = 0;
                    var registro = 0;

                    while (!sr.EndOfStream)
                    {
                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');
                        var tmp = linhax[10].Replace("-", "").Replace(".", "").Trim().PadLeft(8, '0');
                        linhax[10] = tmp.Substring(000, 005) + "-" + tmp.Substring(005, 003);

                        if (linhax[01].Trim().Length < 12)
                            linhax[01] = Convert.ToUInt64(linhax[01]).ToString(@"000\.000\.000\-00");
                        else
                            linhax[01] = Convert.ToUInt64(linhax[01]).ToString(@"00\.000\.000\/0000\-00");

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[10].Replace("-", "").Trim() == "" ? "95096100" : linhax[10].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        registro++;

                        string[] NR = linhax[5].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "DVV";

                        sw.WriteLine(registro.ToString("d8") + "^" +
                                     string.Join("^", linhax) + "^" +
                                     nmCDD.Trim() + "^" +
                                     CIF.Trim() + "^" +
                                     datamatix + "^" +
                                     "São Paulo, " + string.Format("{0:dd 'de' MMMM 'de' yyyy}", DateTime.Now) + "^");

                        carga.WriteLine(linhax[2].Trim() + ";" + //chavepesquisa
                                        linhax[3].Trim() + ";" + //destinatario
                                        linhax[3].Trim() + ";" + //nomekit
                                        linhax[3].Trim() + ";" + //nome
                                        linhax[4].Trim() + ";" + //endereco
                                        linhax[5].Trim() + ";" + //numero
                                        linhax[6].Trim() + ";" + //complemento
                                        linhax[7].Trim() + ";" + //bairro
                                        linhax[8].Trim() + ";" + //cidade
                                        linhax[9].Trim() + ";" + //estado
                                        linhax[10].Trim() + ";" + //cep
                                        linhax[1].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (4.80).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "5" + ";" + //qtdepaginas
                                        "5" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        linhax[2].Trim() + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA DVVDUPLICIDADE" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();

                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch
            {
                return false;
            }

        }

        private static void OrdenarCartasDVVDUPLICIDADE(string Arquivo)
        {
            try
            {
                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                    CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                    StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                    string linha = sr.ReadLine();
                    sw.WriteLine(linha);

                    while (!sr.EndOfStream)
                    {
                        linha = sr.ReadLine();

                        _CarneAtual.DadosCarne = linha;
                        string[] tmp = linha.Split(';');
                        Int32 CEP = 0;
                        Int32.TryParse(tmp[10].Replace("-", "").Trim() == "" ? "95096100" : tmp[10].Replace("-", "").Trim(), out CEP);

                        _CarneAtual.CEP = CEP;
                        _CarneAtual.Parcelas.Add(linha);
                        _Carnes.Add(_CarneAtual);
                        _CarneAtual = new CarneOrdenaCEP();

                    }

                    var lista = _Carnes.OrderBy(x => x.CEP);
                    foreach (var item in lista)
                    {
                        foreach (var parcela in item.Parcelas)
                            sw.WriteLine(parcela);
                    }
                    sw.Dispose();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void processarBACENToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasBACEN(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    //Applicatio/n.Exit();
                }
            }

        }

        private static bool ProcessarCartasBACEN(string Arquivo)
        {
            try
            {
                var seqarq = 0;
                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\BACEN\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCarta(Arquivo, "BACEN");
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;

                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_BACEN", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_BACEN (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_BACEN", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = null;
                StreamWriter carga = new StreamWriter(diretorio + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter relatorio = new StreamWriter(diretorio + "//EPF_RECIBO_GRAFICA.txt", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();
                    SequencialFAC = 0;

                    while (!sr.EndOfStream)
                    {
                        if (SequencialFAC % 1500 == 0)
                        {
                            if (seqarq > 0)
                            {
                                sw.Dispose();
                            }

                            seqarq++;
                            var prearquivo = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(Arquivo)) + "_BACEN_" + seqarq.ToString("d3");
                            NM_ARQUIVO = diretorio + Path.ChangeExtension(prearquivo, "SAI");
                            sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                        }

                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[6].Replace("-", "").Trim() == "" ? "95096100" : linhax[6].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        string[] NR = linhax[3].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "CARTAS BACEN";

                        sw.WriteLine(linha + ";" +
                                     nmCDD.Trim() + ";" +
                                     CIF.Trim() + ";" +
                                     datamatix);

                        relatorio.WriteLine(linhax[0].Trim() + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DateTime.Now) + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DataPostagem) + ";" +
                                            Lote.ToString("d5") + ";");

                        carga.WriteLine(linhax[0].Trim() + ";" + //chavepesquisa
                                        linhax[2].Trim() + ";" + //destinatario
                                        linhax[2].Trim() + ";" + //nomekit
                                        linhax[2].Trim() + ";" + //nome
                                        linhax[3].Trim() + ";" + //endereco
                                        numeroresidencia + ";" + //numero
                                        linhax[4].Trim() + ";" + //complemento
                                        "" + ";" + //bairro
                                        linhax[5].Trim() + ";" + //cidade
                                        "" + ";" + //estado
                                        linhax[6].Trim() + ";" + //cep
                                        linhax[0].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (4.80).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "2" + ";" + //qtdepaginas
                                        "1" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        "" + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA BACEN" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();
                relatorio.Dispose();
                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        private static bool ProcessarCartasConsigna(string Arquivo)
        {
            try
            {
                var seqarq = 0;
                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartaConsignado(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;

                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_DVVDUPLICIDADE", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_DVVDUPLICIDADE (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_DVVDUPLICIDADE", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = null;
                StreamWriter carga = new StreamWriter(diretorio + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter relatorio = new StreamWriter(diretorio + "//EPF_RECIBO_GRAFICA.txt", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();
                    SequencialFAC = 0;

                    while (!sr.EndOfStream)
                    {
                        if (SequencialFAC % 15000 == 0)
                        {
                            if (seqarq > 0)
                            {
                                sw.Dispose();
                            }

                            seqarq++;
                            var prearquivo = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(Arquivo)) + "_" + seqarq.ToString("d3");
                            NM_ARQUIVO = diretorio + Path.ChangeExtension(prearquivo, "SAI");
                            sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                        }

                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[11].Replace("-", "").Trim() == "" ? "95096100" : linhax[11].Replace("-", "").Trim(), out CEP);

                        var tmp = linhax[11].Trim().Replace("-", "").PadLeft(8, '0');
                        linhax[11] = tmp.Substring(00, 05) + "-" + tmp.Substring(05, 03);

                        try
                        {
                            if (Convert.ToUInt64(linhax[04].Trim()).ToString().Length < 12)
                                linhax[04] = Convert.ToUInt64(linhax[04]).ToString(@"000\.000\.000\-00");
                            else
                                linhax[04] = Convert.ToUInt64(linhax[04]).ToString(@"00\.000\.000\/0000\-00");
                        }
                        catch
                        {
                            linhax[04] = linhax[04];
                        }

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        string[] NR = linhax[8].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "CONSIGNADO";

                        sw.WriteLine(string.Join(";", linhax) + ";" +
                                     nmCDD.Trim() + ";" +
                                     CIF.Trim() + ";" +
                                     datamatix + ";" +
                                     "São Paulo, 04 de maio de 2020" + ";");

                        relatorio.WriteLine(linhax[0].Trim() + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DateTime.Now) + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DataPostagem) + ";" +
                                            Lote.ToString("d5") + ";");

                        carga.WriteLine(linhax[4].Trim() + ";" + //chavepesquisa
                                        linhax[5].Trim() + ";" + //destinatario
                                        linhax[5].Trim() + ";" + //nomekit
                                        linhax[5].Trim() + ";" + //nome
                                        linhax[7].Trim() + ";" + //endereco
                                        numeroresidencia + ";" + //numero
                                        linhax[9].Trim() + ";" + //complemento
                                        "" + ";" + //bairro
                                        linhax[10].Trim() + ";" + //cidade
                                        "" + ";" + //estado
                                        linhax[11].Trim() + ";" + //cep
                                        linhax[0].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (4.80).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "5" + ";" + //qtdepaginas
                                        "5" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        "" + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA CONSIGNADO" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();
                relatorio.Dispose();
                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        private static void OrdenarCartaConsignado(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();
                sw.WriteLine(linha);

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');
                    Int32 CEP = 0;
                    Int32.TryParse(tmp[11].Replace("-", "").Trim() == "" ? "95096100" : tmp[11].Replace("-", "").Trim(), out CEP);

                    _CarneAtual.CEP = CEP;
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new CarneOrdenaCEP();

                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }

        private void toolStripMenuItem15_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasConsigna(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }

        }

        private void toolStripMenuItem15_Click_1(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasVeiculos(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }

        }

        private static bool ProcessarCartasVeiculos(string Arquivo)
        {
            try
            {
                var seqarq = 0;
                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartaVeiculos(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;

                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_DVVDUPLICIDADE", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_DVVDUPLICIDADE (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_DVVDUPLICIDADE", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = null;
                StreamWriter carga = new StreamWriter(diretorio + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter relatorio = new StreamWriter(diretorio + "//EPF_RECIBO_GRAFICA.txt", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();
                    SequencialFAC = 0;

                    while (!sr.EndOfStream)
                    {
                        if (SequencialFAC % 15000 == 0)
                        {
                            if (seqarq > 0)
                            {
                                sw.Dispose();
                            }

                            seqarq++;
                            var prearquivo = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(Arquivo)) + "_" + seqarq.ToString("d3");
                            NM_ARQUIVO = diretorio + Path.ChangeExtension(prearquivo, "SAI");
                            sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                        }

                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[9].Replace("-", "").Trim() == "" ? "95096100" : linhax[9].Replace("-", "").Trim(), out CEP);

                        var tmp = linhax[9].Trim().Replace("-", "").PadLeft(8, '0');
                        linhax[9] = tmp.Substring(00, 05) + "-" + tmp.Substring(05, 03);

                        try
                        {
                            if (Convert.ToUInt64(linhax[03].Trim()).ToString().Length < 12)
                                linhax[03] = Convert.ToUInt64(linhax[03]).ToString(@"000\.000\.000\-00");
                            else
                                linhax[03] = Convert.ToUInt64(linhax[03]).ToString(@"00\.000\.000\/0000\-00");
                        }
                        catch
                        {
                            linhax[03] = linhax[03];
                        }

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        string[] NR = linhax[6].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "VEICULOS";

                        sw.WriteLine(SequencialFAC.ToString("d8") + ";" +
                                     string.Join(";", linhax) + ";" +
                                     nmCDD.Trim() + ";" +
                                     CIF.Trim() + ";" +
                                     datamatix + ";" +
                                     "São Paulo, 04 de maio de 2020" + ";");

                        relatorio.WriteLine(linhax[0].Trim() + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DateTime.Now) + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DataPostagem) + ";" +
                                            Lote.ToString("d5") + ";");

                        carga.WriteLine(linhax[0].Trim() + ";" + //chavepesquisa
                                        linhax[1].Trim() + ";" + //destinatario
                                        linhax[1].Trim() + ";" + //nomekit
                                        linhax[1].Trim() + ";" + //nome
                                        linhax[6].Trim() + ";" + //endereco
                                        numeroresidencia + ";" + //numero
                                        "" + ";" + //complemento
                                        "" + ";" + //bairro
                                        linhax[7].Trim() + ";" + //cidade
                                        linhax[7].Trim() + ";" + //estado
                                        linhax[9].Trim() + ";" + //cep
                                        linhax[0].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (4.80).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "5" + ";" + //qtdepaginas
                                        "5" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        "" + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA VEICULOS" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();
                relatorio.Dispose();
                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        private static void OrdenarCartaVeiculos(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();
                sw.WriteLine(linha);

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');
                    Int32 CEP = 0;
                    Int32.TryParse(tmp[9].Replace("-", "").Trim() == "" ? "95096100" : tmp[11].Replace("-", "").Trim(), out CEP);

                    _CarneAtual.CEP = CEP;
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new CarneOrdenaCEP();

                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }

        private void processarToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasCCT(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    // Application.Exit();
                }
            }
        }

        private static bool ProcessarCartasCCT(string Arquivo)
        {
            try
            {
                var seqarq = 0;
                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\CCT\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCarta(Arquivo, "CCT");
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;

                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CCT", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_CCT (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CCT", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = null;
                StreamWriter carga = new StreamWriter(diretorio + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter relatorio = new StreamWriter(diretorio + "//EPF_RECIBO_GRAFICA.txt", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();
                    SequencialFAC = 0;

                    while (!sr.EndOfStream)
                    {
                        if (SequencialFAC % 1500 == 0)
                        {
                            if (seqarq > 0)
                            {
                                sw.Dispose();
                            }

                            seqarq++;
                            var prearquivo = Path.GetFileNameWithoutExtension(Path.GetFileNameWithoutExtension(Arquivo)) + "_CCT_" + seqarq.ToString("d3");
                            NM_ARQUIVO = diretorio + Path.ChangeExtension(prearquivo, "SAI");
                            sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                        }

                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[6].Replace("-", "").Trim() == "" ? "95096100" : linhax[6].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        string[] NR = linhax[3].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "CARTAS CCT";

                        sw.WriteLine(linha + ";" +
                                     nmCDD.Trim() + ";" +
                                     CIF.Trim() + ";" +
                                     datamatix);

                        relatorio.WriteLine(linhax[0].Trim() + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DateTime.Now) + ";" +
                                            string.Format("{0:dd/MM/yyyy}", DataPostagem) + ";" +
                                            Lote.ToString("d5") + ";");

                        carga.WriteLine(linhax[0].Trim() + ";" + //chavepesquisa
                                        linhax[2].Trim() + ";" + //destinatario
                                        linhax[2].Trim() + ";" + //nomekit
                                        linhax[2].Trim() + ";" + //nome
                                        linhax[3].Trim() + ";" + //endereco
                                        numeroresidencia + ";" + //numero
                                        linhax[4].Trim() + ";" + //complemento
                                        "" + ";" + //bairro
                                        linhax[5].Trim() + ";" + //cidade
                                        "" + ";" + //estado
                                        linhax[6].Trim() + ";" + //cep
                                        linhax[0].Trim() + ";" + //cpfcnpj
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (18.80).ToString() + ";" + //peso 4.80 passa para 18.8
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "6" + ";" + //qtdepaginas
                                        "3" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        "" + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA BACEN" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();
                relatorio.Dispose();
                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }

        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {

        }

        private void processarCB7_MenuItem_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasCB7(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    // Application.Exit();
                }
            }
        }

        private void butProcessarCarne_Click(object sender, EventArgs e)
        {
            String dtTmp = DateTime.Now.ToString("ddMMyyyy");
            string sDataPth = @"C:\Drive_d\PrintN\BancoSafra\Carne\PRODUCAO";
            //    C:\Drive_d\PrintN\Banco Safra\Carne\PRODUCAO\entrada
            string sMaskDir = "*.*";
            string sNovoDir;
            string dirEntrada;
            string nmarquivo;
            string linha;
            bool lAndamento = false;
            bool flagZip = false, flagTxt = false, lSucce = false;

            FileInfo fileInfo;

            List<string> arquivos = null;
            List<string> direc = Directory.GetDirectories(@sDataPth + @"\entrada", sMaskDir, SearchOption.TopDirectoryOnly).ToList();

            try
            {


                foreach (string dirent in direc)
                {
                    if (dirent.ToUpper().Contains("_OK") || dirent.ToUpper().Contains("_ZERADO") || dirent.ToUpper().Contains("_RECUSADO") ||
                        dirent.ToUpper().Contains("-OK") || dirent.ToUpper().Contains("-ZERADO") || dirent.ToUpper().Contains("-RECUSADO") ||
                        dirent.ToUpper().Contains("-CANCEL") || dirent.ToUpper().Contains("-ERRO") || dirent.ToUpper().Contains("-INVALIDO") ||
                        dirent.ToUpper().Contains("_CANCEL") || dirent.ToUpper().Contains("_ERRO") || dirent.ToUpper().Contains("_INVALIDO") ||
                        dirent.ToUpper().Contains("-PROCESSADO") || dirent.ToUpper().Contains("-AGUARDA") || dirent.ToUpper().Contains("-TESTE") ||
                        dirent.ToUpper().Contains("-PROCESSADO") || dirent.ToUpper().Contains("-AGUARDA") || dirent.ToUpper().Contains("_TESTE") ||
                        dirent.ToUpper().Contains("PROCESSADO") || dirent.ToUpper().Contains("AGUARDA") || dirent.ToUpper().Contains(" TESTE"))
                    {
                        continue;
                    }



                    string sMask = "*.zip";
                    arquivos = Directory.GetFiles(@dirent, sMask, SearchOption.TopDirectoryOnly).ToList();

                    foreach (string arqmov in arquivos)
                    {
                        fileInfo = new FileInfo(arqmov);
                        dirEntrada = fileInfo.DirectoryName;
                        nmarquivo = fileInfo.Name.Trim();

                        string teste = arquivos[0];

                        if (arquivos.Count > 0)
                        {
                            string zipPath = fileInfo.FullName;
                            using (ZipArchive archive = ZipFile.OpenRead(zipPath))
                            {
                                foreach (ZipArchiveEntry entry in archive.Entries)
                                {
                                    if (File.Exists(dirEntrada + @"\" + entry.FullName))
                                        File.Delete(dirEntrada + @"\" + entry.FullName);
                                }
                            }
                            ZipFile.ExtractToDirectory(@dirEntrada + @"\" + nmarquivo, @dirEntrada);
                            //MessageBox.Show("Arquivo descompactado com sucesso ...");
                            break;
                        }
                    }



                    sMask = "*.txt";
                    arquivos = Directory.GetFiles(@dirent, sMask, SearchOption.TopDirectoryOnly).ToList();
                    if (arquivos.Count > 0)
                    {
                        fileInfo = new FileInfo(arquivos[0]);
                        dirEntrada = fileInfo.DirectoryName;
                        nmarquivo = fileInfo.Name.Trim();
                        lblMsg.Text = "Processando arquivo ... " + nmarquivo;
                        Application.DoEvents();

                        StreamReader sr = new System.IO.StreamReader(@dirEntrada + @"\" + nmarquivo, Encoding.GetEncoding("iso-8859-1"));
                        linha = sr.ReadLine().Substring(9, 3);
                        sr.Close();
                        if (linha == "CDC")
                        {
                            flagTxt = true;
                            dtTmp = DateTime.Now.ToString("HHmmss.fff"); // complemento temporário da pasta ...
                            sNovoDir = dirent + "_" + dtTmp + "_OK";

                            // renomear pasta ja validada
                            Directory.Move(dirent, sNovoDir);

                            fileInfo = new FileInfo(sNovoDir + @"\" + nmarquivo);
                            dirEntrada = fileInfo.DirectoryName;
                            nmarquivo = fileInfo.Name.Trim();

                            //if (File.Exists(dirEntrada + @"\" + nmarquivo))
                            //    MessageBox.Show("Arquivo OK na pasta ...");



                            DataPostagem = this.dateTimePicker1.Value.Date;

                            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
                            if (dialogResult == DialogResult.Yes)
                            {

                                lblMsg.Text = "Processando Arquivo: " + nmarquivo;
                                lSucce = Processar(sNovoDir + @"\" + nmarquivo);

                            }
                            else
                            {
                                return;
                            }

                            sMask = "*.*";
                            arquivos = Directory.GetFiles(sNovoDir + @"\Processado", sMask, SearchOption.TopDirectoryOnly).ToList();
                            string arq, ext, sWfd;
                            string sArqLog;
                            string sResLog;
                            string sLogProc;

                            if (!Directory.Exists(dirEntrada + @"\pnetlog\"))
                                Directory.CreateDirectory(dirEntrada + @"\pnetlog\");

                            sLogProc = @"\pnetlog\processamento.log";

                            FileMode fileMd = FileMode.OpenOrCreate;//FileMode.CreateNew;
                            if (File.Exists(dirEntrada + sLogProc)) fileMd = FileMode.Append;
                            streamLogProc = new FileStream(dirEntrada + sLogProc, fileMd); //, FileShare.ReadWrite  );
                            swLogProc = new StreamWriter(streamLogProc);
                            lAndamento = true;


                            LogProg(swLogProc, 0, false, dirEntrada + " - " + label5.Text, "", "");


                            foreach (string arqmov in arquivos)
                            {
                                fileInfo = new FileInfo(arqmov);
                                dirEntrada = fileInfo.DirectoryName;
                                nmarquivo = fileInfo.Name.Trim();

                                ext = fileInfo.Extension;
                                arq = fileInfo.Name.Replace(ext, "");
                                sArqLog = arq + ".pnet.log";

                                lblMsg.Text = "Limpando Área ... ";
                                Application.DoEvents();


                                if(  nmarquivo.Contains("0081931751.MUTUO_PF001TMP"))
                                {
                                    string teste = "";
                                }

                                // Faz uma limpeza dos arquivos zerados
                                // e apaga
                                if (Verifica_seArquivoEstaVazio(dirEntrada + @"\" + nmarquivo))
                                {
                                    LogProg(swLogProc, 0, false, "Arquivo vazio apagado! ", nmarquivo, "");
                                   // File.Delete(dirEntrada + @"\" + nmarquivo);

                                    continue;
                                }
                                else
                                {

                                    //Valida capas se foi gerado arquivos de MUtuo e Capa desnecessarios
                                    // e apaga
                                    if ((nmarquivo.Contains("MUTUO_") || nmarquivo.Contains("_CAPA")) || nmarquivo.Contains("_CAPA"))
                                    {
                                        FileInfo file = new FileInfo(dirEntrada + @"\" + nmarquivo);
                                        long totbytes = file.Length;

                                        StreamReader sr2 = new System.IO.StreamReader(@dirEntrada + @"\" + nmarquivo, Encoding.GetEncoding("iso-8859-1"));

                                        int cnt = 0;
                                        bool flgVazio = false;
                                        while (!sr2.EndOfStream)
                                        {
                                            linha = sr2.ReadLine();
                                            cnt++;
                                            if (linha.Substring(0, 20) == "99999999110000000000"  && cnt <= 3)
                                            {
                                                flgVazio = true;
                                                break;
                                            }
                                        }
                                        sr2.Close();

                                        if (flgVazio || totbytes < 350)
                                        {
                                            File.Delete(dirEntrada + @"\" + nmarquivo);
                                            LogProg(swLogProc, 0, false, "Arquivo MUTUO ou CAPA vazio apagado! ", nmarquivo, "");
                                            continue;
                                        }

                                    }
                                }

                                // prepara para gerar PRINT NET
                                string lincmd = "";
                                sWfd = "Carne_EMP58_Mutuo14072020.wfd";

                                if (nmarquivo.Contains("_CAPA") || nmarquivo.Contains("_CONTRACAPA") || nmarquivo.Contains("_MIOLO"))
                                {

                                    lincmd += "InspireCLI.exe "
                                        + @"C:\Drive_d\PrintN\BancoSafra\Carne\WFD\" + sWfd
                                        + " -difEntrada \"" + @dirEntrada + @"\" + nmarquivo + "\""
                                        + " -l \"" + dirEntrada + @"\pnetlog\" + sArqLog + "\""
                                        + " -e PDF -o PDF "
                                        + " -f " + "%h05.%e";

                                    lblMsg.Text = "Gerando PNet ... " + nmarquivo;
                                    Application.DoEvents();

                                    LogProg(swLogProc, 0, false, "Gerando PNet ", "CMD:" + lincmd, "");

                                    sResLog = ExecutarCMD("InspireCLI.exe",
                                                        @"C:\Drive_d\PrintN\BancoSafra\Carne\WFD\" + sWfd
                                                       + " -difEntrada \"" + @dirEntrada + @"\" + nmarquivo + "\""
                                                       + " -l \"" + dirEntrada + @"\pnetlog\" + sArqLog + "\""
                                                       + " -e PDF -o PDF "
                                                       + " -f " + "%h05.%e"
                                                    );
                                }

                            }

                            if (lSucce)
                                MessageBox.Show("Arquivo Processado com sucesso.");
                            else
                                MessageBox.Show("Erro ao processar o arquivo");

                            Application.Exit();

                        }
                    }
                }
            }
            catch (Exception err)
            {
                MessageBox.Show("Erro: 001- " + err.Message);
            }
            finally
            {
                if (lAndamento)
                {
                    swLogProc.Close();

                    if (lSucce)
                        lblMsg.Text = "FIM PROCESSAMENTO! ";
                    else
                        lblMsg.Text = "FIM PROCESSAMENTO! VERIFIQUE O QUE HOUVE DE PROBLEMA. ";
                    Application.DoEvents();

                    Thread.Sleep(2000);
                    lblMsg.Text = " ";
                }
            }

        }

        public static string ExecutarCMD(string comando, string parametros)
        {

            using (Process processo = new Process())
            {
                processo.StartInfo.FileName = Environment.GetEnvironmentVariable("comspec");

                // Formata a string para passar como argumento para o cmd.exe
                processo.StartInfo.Arguments = string.Format("/c {0}", comando + " " + parametros);

                processo.StartInfo.RedirectStandardOutput = true;
                processo.StartInfo.UseShellExecute = false;
                processo.StartInfo.CreateNoWindow = true;

                processo.Start();
                // processo.WaitForExit();7

                //  MessageBox.Show("  OK process \r\n", "  ", MessageBoxButtons.OK);

                string saida = processo.StandardOutput.ReadToEnd();
                return saida;
            }
        }
        public void LogProg(StreamWriter pLogProc, int iCnt, bool eLoop, string Msg1, string Msg2, string Msg3)
        {
            string dtLinhaLog = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            string separa = "------------------------------------------------------------------------------------------------------------------------";

            if (Msg1.Length > 0 && Msg2 == "" && Msg3 == "")
            {
                pLogProc.WriteLine("----------------" + dtLinhaLog + "----------------");
                pLogProc.WriteLine(separa);
                pLogProc.WriteLine(" ");
                pLogProc.WriteLine("ARQUIVO: " + Msg1);
                pLogProc.WriteLine(" ");
                pLogProc.WriteLine(separa);
            }
            else
            {
                if (iCnt <= 2 && eLoop)
                {
                    pLogProc.WriteLine("*** " + dtLinhaLog);
                    pLogProc.WriteLine(Msg2);
                    pLogProc.WriteLine(Msg3);
                }
                else // if (iCnt == 0 && !eLoop)
                {
                    pLogProc.WriteLine("* " + dtLinhaLog);
                    pLogProc.WriteLine(Msg2);
                    pLogProc.WriteLine(Msg3);
                }

            }
            pLogProc.Flush();
        }

        private void TarefaLonga(int p)
        {
            for (int i = 0; i <= 10; i++)
            {
                // faz a thread dormir por "p" milissegundos a cada passagem do loop
                Thread.Sleep(p);
                lblValor.Text = "Tarefa%: " + i.ToString() + " comcluída";
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void butProcessaCartas_Click(object sender, EventArgs e)
        {
            try
            {
                // Create a new TripleDES object to generate a key
                // and an initialization vector (IV).
                using (TripleDES TripleDESalg = TripleDES.Create())
                {
                    // Create a string to encrypt.
                    string sData = "Here is some data to encrypt.";
                    string FileName = "CText.enc";

                    // Encrypt text to a file using the file name, key, and IV.
                    EncryptTextToFile(sData, FileName, TripleDESalg.Key, TripleDESalg.IV);

                    // Decrypt the text from a file using the file name, key, and IV.
                    string Final = DecryptTextFromFile(FileName, TripleDESalg.Key, TripleDESalg.IV);

                    // Display the decrypted string to the console.
                    Console.WriteLine(Final);
                    MessageBox.Show(Final);
                }
            }
            catch (Exception e2)
            {
                Console.WriteLine(e2.Message);
            }
        }


        public static void EncryptTextToFile(String Data, String FileName, byte[] Key, byte[] IV)
        {
            try
            {
                // Create or open the specified file.
                using (FileStream fStream = File.Open(FileName, FileMode.OpenOrCreate))
                {

                    // Create a new TripleDES object.
                    using (TripleDES tripleDESalg = TripleDES.Create())
                    {

                        // Create a CryptoStream using the FileStream
                        // and the passed key and initialization vector (IV).
                        using (CryptoStream cStream = new CryptoStream(fStream,
                            tripleDESalg.CreateEncryptor(Key, IV),
                            CryptoStreamMode.Write))
                        {

                            // Create a StreamWriter using the CryptoStream.
                            using (StreamWriter sWriter = new StreamWriter(cStream))
                            {

                                // Write the data to the stream
                                // to encrypt it.
                                sWriter.WriteLine(Data);
                            }
                        }
                    }
                }
            }
            catch (CryptographicException e)
            {
                Console.WriteLine("A Cryptographic error occurred: {0}", e.Message);
            }
            catch (UnauthorizedAccessException e)
            {
                Console.WriteLine("A file access error occurred: {0}", e.Message);
            }
        }

        public static string DecryptTextFromFile(String FileName, byte[] Key, byte[] IV)
        {
            try
            {
                string retVal = "";
                // Create or open the specified file.
                using (FileStream fStream = File.Open(FileName, FileMode.OpenOrCreate))
                {

                    // Create a new TripleDES object.
                    using (TripleDES tripleDESalg = TripleDES.Create())
                    {

                        // Create a CryptoStream using the FileStream
                        // and the passed key and initialization vector (IV).
                        using (CryptoStream cStream = new CryptoStream(fStream,
                            tripleDESalg.CreateDecryptor(Key, IV),
                            CryptoStreamMode.Read))
                        {

                            // Create a StreamReader using the CryptoStream.
                            using (StreamReader sReader = new StreamReader(cStream))
                            {

                                // Read the data from the stream
                                // to decrypt it.
                                retVal = sReader.ReadLine();
                            }
                        }
                    }
                }
                // Return the string.
                return retVal;
            }
            catch (CryptographicException e)
            {
                Console.WriteLine("A Cryptographic error occurred: {0}", e.Message);
                return null;
            }
            catch (UnauthorizedAccessException e)
            {
                Console.WriteLine("A file access error occurred: {0}", e.Message);
                return null;
            }
        }

        private static bool ProcessarCartasDOCLA(string Arquivo)
        {
            try
            {
                NM_ARQUIVO = Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Path.ChangeExtension(Arquivo, "SAI"));

                var diretorio = Path.GetDirectoryName(Arquivo) + "\\Processado\\";
                Directory.CreateDirectory(diretorio);
                OrdenarCartasDOCLA(Arquivo);
                Arquivo = diretorio + Path.GetFileName(Arquivo);

                var nrlote = 0;
                #region Pega o Lote
                OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB + ";Persist Security Info=True;");
                connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CQC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    nrlote = reader.GetInt32(0) + 1;
                    Lote = Convert.ToInt32(nrlote.ToString());
                }
                connection.Close();
                connection.Open();

                command = new OleDbCommand("INSERT INTO FAC_CQC (Lote, DataPostagem) VALUES(" + nrlote + ", '" + string.Format("{0:dd/MM/yyyy}", DataPostagem) + "')", connection);
                reader = command.ExecuteReader();

                command = new OleDbCommand("SELECT MAX(Lote) FROM FAC_CQC", connection);
                reader = command.ExecuteReader();
                while (reader.Read())
                {
                    if (nrlote != reader.GetInt32(0))
                    {
                        MessageBox.Show("Erro ao processar o numero do CIF.");
                        Application.Exit();
                    }
                }
                connection.Close();
                #endregion

                #region Carrega Plano Triagem
                using (StreamReader sr = new StreamReader(Path.GetDirectoryName(ArquivoMDB) + "\\Triagem_CCB.txt", Encoding.GetEncoding("ISO-8859-1")))
                {
                    sr.ReadLine();
                    while (!sr.EndOfStream)
                    {
                        string[] linha = sr.ReadLine().Split('|');
                        ListaTriagem.Add(new Triagem { cepinicial = Convert.ToInt32(linha[0]), cepfinal = Convert.ToInt32(linha[1]), cdd = linha[2], ctc = linha[3] });
                    }
                }
                #endregion

                StreamWriter sw = new StreamWriter(NM_ARQUIVO, false, Encoding.GetEncoding("ISO-8859-1"));
                StreamWriter carga = new StreamWriter(Path.GetDirectoryName(Arquivo) + "//Registros.csv", false, Encoding.GetEncoding("ISO-8859-1"));

                using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
                {
                    var linha = sr.ReadLine();

                    SequencialFAC = 0;
                    var registro = 0;

                    while (!sr.EndOfStream)
                    {
                        SequencialFAC++;
                        linha = sr.ReadLine();
                        var linhax = linha.Split(';');
                        var tmp = linhax[09].Replace("-", "").Replace(".", "").Trim().PadLeft(8, '0');
                        linhax[09] = tmp.Substring(000, 005) + "-" + tmp.Substring(005, 003);

                        //if (linhax[03].Trim().Length < 12)
                        //    linhax[03] = Convert.ToUInt64(linhax[03]).ToString(@"000\.000\.000\-00");
                        //else
                        //    linhax[03] = Convert.ToUInt64(linhax[03]).ToString(@"00\.000\.000\/0000\-00");

                        Int32 CEP = 0;
                        Int32.TryParse(linhax[09].Replace("-", "").Trim() == "" ? "95096100" : linhax[09].Replace("-", "").Trim(), out CEP);

                        string Destino = "3";
                        if (CEP < 10000000)
                            Destino = "1";
                        else if (CEP < 20000000)
                            Destino = "2";

                        var nmCDD = "Não Localizado";
                        var CDD = ListaTriagem.Where(x => CEP >= x.cepinicial && CEP <= x.cepfinal).FirstOrDefault();

                        if (CDD != null)
                            nmCDD = (CDD.cdd.Trim() + " / " + CDD.ctc.Trim()).PadRight(100).Substring(000, 100);

                        var CIF = "7211181591" + Lote.ToString("d5") + SequencialFAC.ToString("d11") + Destino + "0" + string.Format("{0:ddMMyy}", DataPostagem);

                        registro++;

                        string[] NR = linhax[4].Trim().Split(' ');
                        string numeroresidencia = "";
                        int nr = NR.Count();

                        try
                        {
                            for (int m = 0; m < nr; m++)
                            {
                                numeroresidencia = NR[m];
                                int n = 0;
                                var isNumeric = int.TryParse(numeroresidencia, out n);
                                if (isNumeric)
                                {
                                    numeroresidencia = n.ToString();
                                    break;
                                }
                                numeroresidencia = "0";
                            }
                        }
                        catch { }

                        int nx = 0;
                        var isNumericx = int.TryParse(numeroresidencia, out nx);
                        if (!isNumericx)
                            numeroresidencia = "0";

                        string cddestino = "82031";
                        if (Destino == "1")
                            cddestino = "82015";
                        else if (Destino == "2")
                            cddestino = "82023";

                        var datamatix = CEP.ToString("d8") +
                                        numeroresidencia.PadLeft(5, '0').Substring(000, 005) +
                                        "01310930" +
                                        "02150" +
                                        ModCep2D(CEP.ToString("d8")) +
                                        "01" +
                                        CIF +
                                        "0000000000" +
                                        cddestino +
                                        "000000000000000" +
                                        "006422100" +
                                        "|" +
                                        "DOCLA";

                        sw.WriteLine(registro.ToString("d8") + "^" +
                                     string.Join("^", linhax) + "^" +
                                     nmCDD.Trim() + "^" +
                                     CIF.Trim() + "^" +
                                     datamatix + "^" +
                                     "São Paulo, " + string.Format("{0:dd 'de' MMMM 'de' yyyy}", DateTime.Now) + "^");

                        carga.WriteLine(linhax[1].Trim() + ";" + //chavepesquisa
                                        linhax[3].Trim() + ";" + //destinatario
                                        linhax[0].Trim() + ";" + //nomekit
                                        linhax[2].Trim() + ";" + //nome
                                        linhax[3].Trim() + ";" + //endereco
                                        linhax[4].Trim() + ";" + //numero
                                        linhax[5].Trim() + ";" + //complemento
                                        linhax[6].Trim() + ";" + //bairro
                                        linhax[7].Trim() + ";" + //cidade
                                        linhax[8].Trim() + ";" + //estado
                                        linhax[9].Trim() + ";" + //cep
                                        linhax[1].Trim() + ";" + //CONTRATO
                                        CIF + ";" + //rastreio
                                        CIF.Substring(26, 01) + ";" + //localidade
                                        "FAC" + ";" + //tipopostagem
                                        "Processado" + ";" + //status
                                        "" + ";" + //referencia
                                        Lote.ToString() + ";" + //os
                                        SequencialFAC.ToString("d8").Substring(000, 003) + ";" + //loteid
                                        SequencialFAC.ToString("d8") + ";" + //audit
                                        (9.60).ToString() + ";" + //peso
                                        "1" + ";" + //paginainicial
                                        "2" + ";" + //paginafinal
                                        "5" + ";" + //qtdepaginas
                                        "5" + ";" + //qtdefolhas
                                        "P" + ";" + //familia
                                        "" + ";" + //codinsumo01
                                        "" + ";" + //descricao01
                                        "" + ";" + //codinsumo02
                                        "" + ";" + //descricao02
                                        "" + ";" + //codinsumo03
                                        "" + ";" + //descricao03
                                        "" + ";" + //codinsumo04
                                        "" + ";" + //descricao04
                                        "" + ";" + //codinsumo05
                                        "" + ";" + //descricao05
                                        "" + ";" + //crt
                                        "" + ";" + //etq
                                        "" + ";" + //crn
                                        "" + ";" + //manual
                                        "" + ";" + //codigomanual
                                        "" + ";" + //auditpostagem
                                        "" + ";" + //auditkit
                                        "" + ";" + //auditado
                                        "" + ";" + //fechado
                                        "" + ";" + //arquivoentradaid
                                        "" + ";" + //datarecepcao
                                        "" + ";" + //datapostagem
                                        "" + ";" + //dataentrega
                                        "" + ";" + //tipokit
                                        "" + ";" + //leitura
                                        "" + ";" + //email
                                        linhax[1].Trim() + ";" + //contrato
                                        "" + ";" + //nomepdf
                                        "" + ";" + //telefone
                                        "" + ";" + //pathsms
                                        "" + ";" + //pathemail
                                        "" + ";" + //pdfgerado
                                        "" + ";" + //codigocliente
                                        "CARTA DOCLA" + ";"   //produto
                                        );

                    }
                }
                sw.Dispose();
                carga.Dispose();

                #region Grava Dados
                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ArquivoMDB);
                con.Open();

                string sqlstr = "INSERT INTO Registros SELECT * FROM [Text;HDR=NO;DATABASE=" + Path.GetDirectoryName(Arquivo) + "\\].Registros.csv";

                if (ArquivoVazio(Path.GetDirectoryName(Arquivo) + "//Registros.csv"))
                {
                    Registros(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                    OleDbCommand olecom = new OleDbCommand(sqlstr, con);
                    olecom.ExecuteNonQuery();
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Registros.csv");
                    File.Delete(Path.GetDirectoryName(Arquivo) + "//Schema.ini");
                }

                #endregion
                return true;
            }
            catch
            {
                return false;
            }

        }

        private void ProcessarCartaDOCLA_Click(object sender, EventArgs e)
        {
            DataPostagem = this.dateTimePicker1.Value.Date;

            DialogResult dialogResult = MessageBox.Show("Confirma Data de Postagen para " + string.Format("{0:dd/MM/yyyy}", DataPostagem), "Data de Postagem", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = "Processando Arquivo: " + Path.GetFileName(openFileDialog1.FileName);
                    if (ProcessarCartasDOCLA(openFileDialog1.FileName))
                        MessageBox.Show("Arquivo Processado com sucesso.");
                    else
                        MessageBox.Show("Erro ao processar o arquivo");

                    Application.Exit();
                }
            }
        }

        private static void OrdenarCartasDOCLA(string Arquivo)
        {
            using (StreamReader sr = new StreamReader(Arquivo, Encoding.GetEncoding("ISO-8859-1")))
            {
                List<CarneOrdenaCEP> _Carnes = new List<CarneOrdenaCEP>();
                CarneOrdenaCEP _CarneAtual = new CarneOrdenaCEP();

                StreamWriter sw = new StreamWriter(Path.GetDirectoryName(Arquivo) + "\\Processado\\" + Path.GetFileName(Arquivo), false, Encoding.GetEncoding("ISO-8859-1"));
                string linha = sr.ReadLine();
                sw.WriteLine(linha);

                while (!sr.EndOfStream)
                {
                    linha = sr.ReadLine();

                    _CarneAtual.DadosCarne = linha;
                    string[] tmp = linha.Split(';');
                    Int32 CEP = 0;
                    Int32.TryParse(tmp[09].Replace("-", "").Trim() == "" ? "95096100" : tmp[09].Replace("-", "").Trim(), out CEP);

                    _CarneAtual.CEP = CEP;
                    _CarneAtual.Parcelas.Add(linha);
                    _Carnes.Add(_CarneAtual);
                    _CarneAtual = new CarneOrdenaCEP();

                }

                var lista = _Carnes.OrderBy(x => x.CEP);
                foreach (var item in lista)
                {
                    foreach (var parcela in item.Parcelas)
                        sw.WriteLine(parcela);
                }
                sw.Dispose();

            }
        }



    }


}

