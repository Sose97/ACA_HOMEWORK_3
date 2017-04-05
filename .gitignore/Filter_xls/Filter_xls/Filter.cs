using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Filter_xls
{
    class Filter
    {
        public Func<Student,bool> NameFilter(string name)
        {
            return (s => s.Name.Contains(name));
        }

        public Func<Student,bool> SureNameFilter(string surname)
        {
            return (s => s.Surname.Contains(surname));
        }

        public Func<Student, bool> BirthDayFilter(int low, int up)
        {
            return (s => s.BirthData >= low && s.BirthData <= up);
        }

        public Func<Student, bool> KnownLangFilter(int low, int up)
        {
            return (s => s.KnownLang >= low && s.KnownLang <= up);
        }

        public Func<Student, bool> PhoneNumberFilter(string phone)
        {
            return (s => s.PhoneNumber.Contains(phone));
        }
    }
}
