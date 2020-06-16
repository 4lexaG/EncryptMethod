using Sample;
using System;
using System.Security.Cryptography;
using System.Text;

namespace EncryptDecrypt
{
    public class EncryptMethod
    {
        enum TransformType { ENCRYPT = 0, DECRYPT = 1 }
        UTF8Encoding _enc;
        RijndaelManaged _rcipher;
        byte[] _key, _pwd, _ivBytes, _iv;

        private string[] Transform(string file, string[] data)
        {
            String iv = "";
            string key = "";

            var application = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document document = application.Documents.Open(file);

            int count = document.Words.Count;
            string s2 = "";
            for (int i = 1; i <= count; i++)
            {
                string text = document.Words[i].Text;
                s2 = s2 + text;
            }
            String cypherText = "";
                       
            if (data[0] == null)
            {
                iv = CryptLib.GenerateRandomIV(16);
                Random random = new Random();

                key = getHashSha256(random.Next(6, 8).ToString(), 31);
                data[0] = iv;
                data[1] = key;
            }
            cypherText = Encrypt(s2, data[1], data[0]);

            document.Content.Text = cypherText;
            document.Close(true);
            application.Quit(true);
            return data;
        }

        private String Encrypt(string _inputText, string _encryptionKey, string _initVector)
        {

            string _out = "";
            _pwd = Encoding.UTF8.GetBytes(_encryptionKey);
            _ivBytes = Encoding.UTF8.GetBytes(_initVector);

            int len = _pwd.Length;
            if (len > _key.Length)
            {
                len = _key.Length;
            }
            int ivLenth = _ivBytes.Length;
            if (ivLenth > _iv.Length)
            {
                ivLenth = _iv.Length;
            }

            Array.Copy(_pwd, _key, len);
            Array.Copy(_ivBytes, _iv, ivLenth);
            _rcipher.Key = _key;
            _rcipher.IV = _iv;
            
            byte[] plainText = _rcipher.CreateEncryptor().TransformFinalBlock(_enc.GetBytes(_inputText), 0, _inputText.Length);
            _out = Convert.ToBase64String(plainText);
            
            _rcipher.Clear();
            _rcipher.Dispose();
            return _out;
        }
        private string getHashSha256(string text, int length)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(text);
            SHA256Managed hashstring = new SHA256Managed();
            byte[] hash = hashstring.ComputeHash(bytes);
            string hashString = string.Empty;
            foreach (byte x in hash)
            {
                hashString += String.Format("{0:x2}", x); 
            }
            if (length > hashString.Length)
                return hashString;
            else
                return hashString.Substring(0, length);
        }

    }
}