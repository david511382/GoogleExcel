using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.OpenSsl;
using Org.BouncyCastle.Security;
using System;
using System.IO;
using System.Text;

namespace GoogleExcel
{
    internal static class ConvertUtil
    {
        public static string Base64(this string data)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(data);
            return bytes.Base64();
        }

        public static string Base64(this byte[] bytes)
        {
            return Convert.ToBase64String(bytes);
        }

        public static byte[] Sha256WithRSA(string data, string privateKey)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(data);
            ISigner signer = SignerUtilities.GetSigner("SHA256withRSA");

            signer.Init(true, pkcs1(privateKey));

            signer.BlockUpdate(bytes, 0, bytes.Length);
            return signer.GenerateSignature();
        }

        private static AsymmetricKeyParameter pkcs1(string key)
        {
            using (StringReader txtreader = new StringReader(key))
            {
                object t = new PemReader(txtreader).ReadObject();
                AsymmetricKeyParameter keyPair = (AsymmetricKeyParameter)t;
                return keyPair;
            }
        }
    }
}
