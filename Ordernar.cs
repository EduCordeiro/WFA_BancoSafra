using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace WindowsFormsApplication1
{
    public class Ordenar
    {
        public string Header { get; set; }
        public string Trailer { get; set; }
        public string Chave01 { get; set; }
        public int Chave01Inicio { get; set; }
        public int Chave01Tamanho { get; set; }

        //public CEP CEP { get; set; }
        public string CEPxx { get; set; }

        public bool Ordenar_Arquivo(string Arquivo)
        {
            try
            {
                #region
                using (StreamReader sr = new StreamReader(Arquivo, Encoding.Default))
                {
                    List<Ordenacao> ListaOrdenacao = new List<Ordenacao>();
                    string _Header = null;
                    string _Trailler = null;
                    Int64 TMLinha = 0;
                    Int64 PosiIni = 0;
                    int TamanhoBloco = 0;
                    int Objeto = 0;
                    int Paginas = 0;

                    #region Classifica o arquivo e insere o CDD
                    while (!sr.EndOfStream)
                    {
                        string linha = sr.ReadLine();

                        try
                        {
                            if (linha.Substring(Chave01Inicio, Chave01Tamanho) == Header)
                                _Header = linha;
                            else if (linha.Substring(Chave01Inicio, Chave01Tamanho) == Trailer)
                                _Trailler = linha;
                            else if (linha.Substring(Chave01Inicio, Chave01Tamanho) == Chave01)
                            {
                                if (Objeto > 0)
                                {
                                    ListaOrdenacao.Add(new Ordenacao()
                                    {
                                        CEP = CEPxx, //CEP = CEP.Cep,
                                        INDEX = 1,//Convert.ToInt64((string.IsNullOrEmpty(CEP.CodigoCDD) ? "9999" : CEP.CodigoCDD.Trim())),
                                        Inicio = PosiIni,
                                        Paginas = Paginas,
                                        Tamanho = TamanhoBloco
                                    });
                                }

                                Paginas = 0;

                                #region Validação de CEP
                                try
                                {
                                    //string[] campos = linha.Split(new[] { "|&" }, StringSplitOptions.None);
                                    //string CEPx = campos[8].Trim().Replace("-", "");
                                    //CEPxx = CEPx;//new CEP(CEPx, CustomEnum.CentroDeImpressao.Alphaville, CustomEnum.CodigoDRPostagem.SaoPaulo) { Cep = CEPx };
                                    //Paginas = Convert.ToInt16(campos[16]);
                                    CEPxx = linha.Substring(857, 001);
                                    Paginas = 2;
                                }
                                catch
                                {
                                    Console.WriteLine("ex.Message");
                                }
                                #endregion

                                PosiIni = TMLinha;
                                TamanhoBloco = 0;
                                Objeto++;
                            }

                            TMLinha += linha.Length + 2;
                            TamanhoBloco += linha.Length + 2;
                        }
                        catch
                        {
                            Console.WriteLine("ex.Message");
                        }
                    }
                    #endregion

                    //Ultimo objeto do arquivo
                    if (Objeto > 0)
                    {
                        ListaOrdenacao.Add(new Ordenacao()
                        {
                            CEP = CEPxx, //CEP.Cep,
                            INDEX = 1, //Convert.ToInt16((CEP.CepValido == true ? CEP.CodigoDestino.ToString() : "9999")),
                            Inicio = PosiIni,
                            Paginas = Paginas,
                            Tamanho = TamanhoBloco
                        });
                    }

                    #region Grava no arquivo de saida
                    if (ListaOrdenacao.Count > 0)
                    {
                        //Classifica o arquivo
                        try
                        {
                            FileStream Saida = new FileStream(Arquivo + ".ORD", FileMode.Create, FileAccess.Write);

                            var ordena = from s in ListaOrdenacao orderby s.Paginas, s.INDEX, s.CEP select s;
                            //var ordena = ListaOrdenacao.OrderBy(c => c.Paginas).ThenBy(d => d.CEP).ToList();

                            using (FileStream Entrada = new FileStream(Arquivo, FileMode.Open, FileAccess.Read))
                            {
                                foreach (Ordenacao linha in ordena)
                                {
                                    Entrada.Position = linha.Inicio;
                                    while (Entrada.Length > Entrada.Position)
                                    {
                                        Int64 position = (Int64)Entrada.Position;
                                        int TMBloco = linha.Tamanho;

                                        //Pega o tamanho do bloco
                                        byte[] byts = new byte[TMBloco];
                                        Entrada.Read(byts, 0, TMBloco);
                                        Saida.Write(byts, 0, TMBloco);
                                        break;
                                    }
                                }
                            }
                            Saida.Dispose();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                    #endregion
                }
                #endregion
                //File.Delete(Arquivo);
                return true;
            }
            catch
            {
                return false;
            }

        }

        private class Ordenacao
        {
            public Int64 Inicio { get; set; }
            public int Tamanho { get; set; }
            public string CEP { get; set; }
            public int Paginas { get; set; }
            public Int64 INDEX { get; set; }
        }
    }
}
