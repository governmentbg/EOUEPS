using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPayments.Common.Helpers
{
    public class EgnHelper
    {
        public enum Gender
        {
            /// <summary>
            /// Стойност по подразбиране: пола на лицето не е инициализиран
            /// </summary>
            None,
            /// <summary>
            /// Лицето е мъж
            /// </summary>
            Male,
            /// <summary>
            /// Лицето е жена
            /// </summary>
            Female
        }

        private int[] coef = { 2, 4, 8, 5, 10, 9, 7, 3, 6 };
        private DateTime birthDate;
        private Gender sex;

        public string EGNError;
        public bool IsValid()
        {
            return string.IsNullOrEmpty(EGNError);
        }

        public EgnHelper(string egn)
        {
            if (string.IsNullOrWhiteSpace(egn) || egn.Length != 10)
            {
                EGNError = "Invalid EGN length";
                return;
            }

            for (int i = 0; i < 10; i++)
                if (!Char.IsDigit(egn, i))
                {
                    EGNError = "EGN must contain digits only";
                    return;
                }

            int yy = Int32.Parse(egn.Substring(0, 2));
            int mm = Int32.Parse(egn.Substring(2, 2));
            int dd = Int32.Parse(egn.Substring(4, 2));

            if (mm >= 21 && mm <= 32)
            {
                mm -= 20;
                yy += 1800;
            }
            else if (mm >= 41 && mm <= 52)
            {
                mm -= 40;
                yy += 2000;
            }
            else
                yy += 1900;

            try
            {
                birthDate = new DateTime(yy, mm, dd);
            }
            catch
            {
                EGNError = "Invalid date in EGN";
                return;
            }

            if (Convert.ToInt32(egn.Substring(8, 1)) % 2 == 0)
                sex = Gender.Male;
            else
                sex = Gender.Female;

            int sum = 0;
            for (int i = 0; i < 9; i++)
                sum += (Int32.Parse(egn.Substring(i, 1)) * coef[i]);

            int rem = sum % 11;
            if (rem == 10)
                rem = 0;

            if (rem != Int32.Parse(egn.Substring(9, 1)))
            {
                EGNError = "Invalid EGN checksum";
                return;
            }
        }

        public static DateTime DateFromEGNInternal(string egn)
        {
            EgnHelper e = new EgnHelper(egn);
            if (!e.IsValid())
            {
                return DateTime.MinValue;
            }
            else
            {
                return e.BirthDate;
            }
        }

        /// <summary>
        /// Датата на раждане на лицето със съответното ЕГН.
        /// </summary>
        public DateTime BirthDate
        {
            get
            {
                return birthDate;
            }
        }

        /// <summary>
        /// Полът на лицето със съответното ЕГН.
        /// </summary>
        public Gender Sex
        {
            get
            {
                return sex;
            }
        }

        public static DateTime GetBirthDateFromEgn(string egn)
        {
            EgnHelper e = new EgnHelper(egn);
            if (!e.IsValid())
            {
                return DateTime.MinValue;
            }
            else
            {
                return e.BirthDate;
            }
        }
    }
}
