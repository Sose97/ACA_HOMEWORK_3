using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Filter_xls
{
    class Program
    {
        static void Main(string[] args)
        {
            List<Student> students=new List<Student>();
            FilterValues filterValues = new FilterValues();
            try
            {
                students = Read_From_Excel.getStudentList(@"C:\Users\Sose\Documents\visual studio 2015\Projects\Filter_xls\Filter_xls\Students");
            }catch (Exception)
            {
                Console.WriteLine("Don't find xls file");
            }

            try
            {
                filterValues = Read_From_Excel.getFilterList(@"C:\Users\Sose\Documents\visual studio 2015\Projects\Filter_xls\Filter_xls\Filters");
            }catch(Exception)
            {
                Console.WriteLine("Don't find xls file");
            }

            Filter filter = new Filter();
            List<Student> filteredName = students.Where(filter.NameFilter(filterValues.NameCompare)).ToList();

            Console.WriteLine("1. Name filter");
            foreach (Student item in filteredName)
            {
                Console.WriteLine(item);
            }

            List<Student> filteredSureName = students.Where(filter.SureNameFilter(filterValues.SurnameCompare)).ToList();

            Console.WriteLine("2. Surname filter");
            foreach (Student item in filteredSureName)
            {
                Console.WriteLine(item);
            }

            List<Student> filteredBirthDay = students.Where(filter.BirthDayFilter(filterValues.BirhtDataFrom, filterValues.BirthDataTo)).ToList();

            Console.WriteLine("3. BirthDay filter");
            foreach (Student item in filteredBirthDay)
            {
                Console.WriteLine(item);
            }

            List<Student> filteredKnownLanguage = students.Where(filter.KnownLangFilter(filterValues.KnownLangFrom, filterValues.KnownLangTo)).ToList();

            Console.WriteLine("4. Known languages count filter");
            foreach (Student item in filteredKnownLanguage)
            {
                Console.WriteLine(item);
            }

            List<Student> filteredPhoneNumber = students.Where(filter.PhoneNumberFilter(filterValues.PhoneNumberCompare)).ToList();

            Console.WriteLine("5. PhoneNaumber filter");
            foreach (Student item in filteredPhoneNumber)
            {
                Console.WriteLine(item);
            }

            Console.WriteLine("Final filter \n The students from Students.xls who satisfy to the conditions from Filters.xls file");
            List<Student> filteredStudent = (((filteredName.Intersect(filteredSureName).ToList()).
                Intersect(filteredBirthDay).ToList()).Intersect(filteredKnownLanguage).ToList()).
                Intersect(filteredPhoneNumber).ToList();
            foreach (Student item in filteredStudent)
            {
                Console.WriteLine(item);
            }
            
        }
    }
}
