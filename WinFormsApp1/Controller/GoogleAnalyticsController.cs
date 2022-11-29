using Google.Authenticator;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinFormsApp1.Model;

namespace WinFormsApp1.Controller
{
    public class GoogleAnalyticsController
    {
        public string otp { get; set; }

        public void CreateSecretKeyAndQrCode(Accounts accounts,out SetupCode code ,out string message)
        {
            message = "";
            code = null;
            try
            {
                TwoFactorAuthenticator twoFactor = new TwoFactorAuthenticator();
                SetupCode query = twoFactor.GenerateSetupCode(accounts.Account, accounts.Account, accounts.Password, false, 3);
                //Todo QRCode圖片從記憶體轉到畫面上
                code = query;
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
        }
        public bool ValidateGoogleAuthCode(Accounts accounts)
        {
            var isRight = false;
            TwoFactorAuthenticator tfA = new TwoFactorAuthenticator();
            isRight = tfA.ValidateTwoFactorPIN(accounts.Password, accounts.otp,TimeSpan.FromSeconds(30));//密碼有效時間
            return isRight;
        }



    }
}
