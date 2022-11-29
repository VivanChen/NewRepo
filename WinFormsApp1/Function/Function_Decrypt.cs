using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace WinFormsApp1.Function
{
    public class Function_Decrypt
    {
        /// <summary>
        /// 默认密钥向量
        /// </summary>
        private static byte[] keys = { 0x12, 0x34, 0x56, 0x78, 0x90, 0xAB, 0xCD, 0xEF };

        /// <summary>
        /// 加密字串
        /// </summary>
        public static string EncryptKey = "12345678";

        /// <summary> 
        /// DES加密字符串 
        /// </summary> 
        /// <param name="encryptString">待加密的字符串</param> 
        /// <returns>加密成功返回加密后的字符串，失败返回源串</returns> 
        public static string EncryptDES(string encryptString)
        {
            return EncryptDES(encryptString, EncryptKey);
        }

        /// <summary> 
        /// DES加密字符串 
        /// </summary> 
        /// <param name="encryptString">待加密的字符串</param> 
        /// <param name="encryptKey">加密密钥,要求为8位</param> 
        /// <returns>加密成功返回加密后的字符串，失败返回源串</returns> 
        public static string EncryptDES(string encryptString, string encryptKey)
        {
            try
            {
                byte[] rgbKey = Encoding.UTF8.GetBytes(encryptKey.Substring(0, 8));
                byte[] rgbIV = keys;
                byte[] inputByteArray = Encoding.UTF8.GetBytes(encryptString);
                DESCryptoServiceProvider desCSP = new DESCryptoServiceProvider();
                MemoryStream memStream = new MemoryStream();
                CryptoStream cryptStream = new CryptoStream(memStream, desCSP.CreateEncryptor(rgbKey, rgbIV), CryptoStreamMode.Write);
                cryptStream.Write(inputByteArray, 0, inputByteArray.Length);
                cryptStream.FlushFinalBlock();
                return Convert.ToBase64String(memStream.ToArray());
            }
            catch
            {
                return encryptString;
            }
        }

        /// <summary>
        /// 解密字串
        /// </summary>
        /// <param name="decryptString">Descrypt String</param>
        /// <returns></returns>
        public static string DecryptDES(string decryptString)
        {
            return DecryptDES(decryptString, EncryptKey);
        }

        /// <summary>
        /// 解密字串.
        /// </summary>
        /// <param name="decryptString">The decrypt string.</param>
        /// <param name="decryptKey">The decrypt key.</param>
        /// <returns></returns>
        public static string DecryptDES(string decryptString, string decryptKey)
        {
            try
            {
                byte[] rgbKey = Encoding.UTF8.GetBytes(decryptKey);
                byte[] rgbIV = keys;
                byte[] inputByteArray = Convert.FromBase64String(decryptString);
                DESCryptoServiceProvider desCSP = new DESCryptoServiceProvider();
                MemoryStream memStream = new MemoryStream();
                CryptoStream cryptoStream = new CryptoStream(memStream, desCSP.CreateDecryptor(rgbKey, rgbIV), CryptoStreamMode.Write);
                cryptoStream.Write(inputByteArray, 0, inputByteArray.Length);
                cryptoStream.FlushFinalBlock();
                return Encoding.UTF8.GetString(memStream.ToArray());
            }
            catch
            {
                return decryptString;
            }
        }
    }

}
