using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinFormsApp1.Model;

namespace WinFormsApp1.Controller
{
    public class AccountsController:Accounts
    {
        string account = "VIVAN";
        string password = "123456";
        string Otp = "";
        public bool Validate_account(Accounts accounts,out string message)
        {
            bool validate = false;
            message = "";
            try
            {             
                //讀取資料庫取得帳號密碼
                //Todo
                if (accounts.Account.Trim() == account &&
                    accounts.Password.Trim() == password)
                {
                    validate = true;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
            }
            return validate;

        }
    }
}
