using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public static class Criptografador
    {
        private static SymmetricAlgorithm ObterTripleDesProvider()
        {
            return new TripleDESCryptoServiceProvider
            {
                KeySize = 192,
                BlockSize = 64,
                Mode = CipherMode.CBC,
                Padding = PaddingMode.Zeros
            };
        }

        private static void TransformarStream(Stream inStream, Stream outStream, ICryptoTransform transformador, int offSet = 0)
        {
            inStream.Seek(offSet, SeekOrigin.Begin);

            FileStream fsread = new FileStream(@"C:\Temp\Teste\FINA_CONS_CONTRATOS1x.ZIP", FileMode.CreateNew);

            using (var cryptoStream = new CryptoStream(inStream, transformador, CryptoStreamMode.Read))
            {
                //cryptoStream.CopyTo(outStream);

                int data;

                while ((data = cryptoStream.ReadByte()) != -1)
                {
                    fsread.WriteByte((byte)data);
                }

               
            }
            fsread.Close();
            MessageBox.Show("Arquivo descriptografado em ");
        }

        public static void Descriptografar(Stream inStream, Stream outStream, byte[] chave, int offSet = 0 )
        {
            var vetorInicializacao = chave.Take(8).ToArray();

            using (var desProvider = ObterTripleDesProvider())
            using (var transformador = desProvider.CreateDecryptor(chave, vetorInicializacao))
            {
                TransformarStream(inStream, outStream, transformador, offSet);
            }
        }

        public static byte[] Criptografar(Stream inStream, Stream outStream, int offSet = 0)
        {
            using (var desProvider = ObterTripleDesProvider())
            {
                // A CIP solicitou que o IV seja configurado com os 8 primeiros digitos da chave
                desProvider.IV = desProvider.Key.Take(8).ToArray();

                using (var transformador = desProvider.CreateEncryptor())
                {
                    TransformarStream(inStream, outStream, transformador, offSet);
                }

                return desProvider.Key;
            }
        }
    }
}
