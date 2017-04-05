using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Filter_xls
{
    class Student
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public int BirthData { get; set; }
        public int KnownLang { get; set; }
        public string PhoneNumber { get; set; }

        public Student(string name,string surname,int bd,int kl,string pn)
        {
            this.Name = name;
            this.Surname = surname;
            this.BirthData = bd;
            this.KnownLang = kl;
            this.PhoneNumber = pn;
        }

        public override string ToString()
        {
            return String.Format(" Name:{0} \n Surname:{1} \n Year Of Birth:{2} \n Number Of Known Programming Languages:{3} \n Phone number:{4} \n\n",
                Name, Surname, BirthData, KnownLang, PhoneNumber);
        }

    }
}
