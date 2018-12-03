using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GenerateurGardes
{
    class Program
    {
        public static int nbAnnees = 10;
        public static DateTime dateDebut;
        public static List<string> week;
        public static List<string> weekend;
        public static Dictionary<string, string> permutations;
        public static List<string> roulementY;
        public static List<string> roulementTotal;
        public static int startYear = 2018;

        public static DateTime currentDate;

        public static void Main(string[] args)
        {
            week = new List<string>();
            weekend = new List<string>();
            permutations = new Dictionary<string, string>();
            roulementY = new List<string>();
            roulementTotal = new List<string>();

            week.Add("Perdicaro");
            week.Add("Pichot");
            week.Add("Delsaut");
            week.Add("Devendeville");
            week.Add("Molmy");
            week.Add("Galerneau");

            weekend.Add("Delsaut");
            weekend.Add("Devendeville");
            weekend.Add("Levecq");
            weekend.Add("Molmy");
            weekend.Add("Galerneau");
            weekend.Add("Lepez");
            weekend.Add("Perdicaro");
            weekend.Add("Pichot");
            weekend.Add("Meresse");

            permutations.Add("Delsaut", "Devendeville");
            permutations.Add("Devendeville", "Molmy");
            permutations.Add("Levecq", "Delsaut");
            permutations.Add("Molmy", "Levecq");
            permutations.Add("Galerneau", "Lepez");
            permutations.Add("Lepez", "Galerneau");

            //Initialisation du calcul pour l'année 1 (2018)
            int startWeek = 0;
            int startWeekend = 8;
            DateTimeFormatInfo dfi = DateTimeFormatInfo.CurrentInfo;
            Calendar cal = dfi.Calendar;
            for (int i = 0; i < nbAnnees; i++)
            {
                //L'annee commence par un jour de semaine ou de weeekend ?
                //Si elle commence par un weekend, on commence par le weekend !
                int currentWeek = initYear(startYear + i, startWeek, startWeekend);

                //Et on continue avec l'année entière
                int[] yearEnds = generateYear(startYear + i, startWeek, startWeekend, currentWeek, cal, dfi);

                //Initialisation du calcul pour l'année suivante
                startWeek = yearEnds[0];
                startWeekend = yearEnds[1];
            }

            //Permutations
            managePermutations(roulementTotal, permutations);

            //Ecriture fichier
            int cpt = 0;
            if (System.IO.File.Exists(@"C:\Users\Pierre\Desktop\gardes.txt"))
            {
                System.IO.File.Delete(@"C:\Users\Pierre\Desktop\gardes.txt");
            }

            for (int j = 0; j < roulementTotal.Count(); j = j + 2)
            {
                using (System.IO.StreamWriter file =
                new System.IO.StreamWriter(@"C:\Users\Pierre\Desktop\gardes.txt", true))
                {
                    if (j % 53 == 0)
                    {
                        file.WriteLine("Année : " + (startYear + cpt));
                        cpt++;
                    }
                    file.WriteLine("Semaine " + ((j / 2) % 53 + 1) + " : " + roulementTotal.ElementAt(j));
                    file.WriteLine("Weekend " + ((j / 2) % 53 + 1) + ": " + roulementTotal.ElementAt(j + 1));
                }
            }

            //Création du fichier Excel
            string path = @"C: \Users\Pierre\Desktop\Gardes.xls";
            if (!File.Exists(path))
            {
                File.Create(path);
            }
            else
            {
                File.Delete(path);

            }

           // ResultExcelFile rFile = new ResultExcelFile(path);

        }

        public static int initYear(int year, int startWeek, int startWeekend)
        {
            roulementY = new List<string>();
            dateDebut = new DateTime(year, 1, 1);
            int currentWeek = 1;
            currentDate = dateDebut;

            //L'annee commence par un jour de semaine ou de weeekend ?
            //Si elle commence par un weekend, on commence par le weekend !
            if (currentDate.DayOfWeek.Equals(DayOfWeek.Saturday) || currentDate.DayOfWeek.Equals(DayOfWeek.Sunday))
            {
                //On définie le premier weekend
                int selected = (startWeekend + (currentWeek - 1)) % weekend.Count();
                roulementY.Add(weekend.ElementAt(selected));
                //et on passe direct à la semaine 2 !
                currentWeek = 2;
            }

            return currentWeek;
        }

        public static int[] generateYear(int year, int startWeek, int startWeekend, int currentWeek, Calendar cal, DateTimeFormatInfo dfi)
        {
            int[] yearsEnds = new int[2];
            int selected;

            //Et on continue avec l'année entière
            for (int i = 1; i < 366; i = i + 7)
            {
                //Détermine les semaines
                selected = (startWeek + (currentWeek - 1)) % week.Count();
                roulementY.Add(week.ElementAt(selected));
                yearsEnds[0] = selected + 1;

                //Détermine les weekends
                selected = (startWeekend + (currentWeek - 1)) % weekend.Count();
                roulementY.Add(weekend.ElementAt(selected));
                yearsEnds[1] = selected + 1;

                currentDate = currentDate.AddDays(7);
                currentWeek = cal.GetWeekOfYear(currentDate, dfi.CalendarWeekRule, dfi.FirstDayOfWeek);
            }

            roulementTotal.AddRange(roulementY);

            return yearsEnds;
        }

        static void managePermutations(List<string> roulement, Dictionary<string, string> permutations)
        {
            for (int i = 0; i < roulement.Count(); i = i + 2)
            {
                if (i > 0)
                {
                    //Permutation weekend / semaine consécutifs
                    if (roulement.ElementAt(i).Equals(roulement.ElementAt(i + 1)))
                    {
                        roulement.Insert(i + 1, permutations[roulement.ElementAt(i)]);
                        roulement.RemoveAt(i + 2);
                    }
                    if (roulement.ElementAt(i).Equals(roulement.ElementAt(i - 1)))
                    {
                        roulement.Insert(i - 1, permutations[roulement.ElementAt(i)]);
                        roulement.RemoveAt(i);
                    }
                }

                //Permutation weekend / semaine non consécutifs (deux weekends de garde consécutifs)
                if (i > 2 && (i + 3) < roulement.Count())
                {
                    if (roulement.ElementAt(i + 1).Equals(roulement.ElementAt(i + 3)))
                    {
                        roulement.Insert(i + 3, permutations[roulement.ElementAt(i + 1)]);
                        roulement.RemoveAt(i + 4);
                    }
                    if (roulement.ElementAt(i - 1).Equals(roulement.ElementAt(i - 3)))
                    {
                        roulement.Insert(i - 3, permutations[roulement.ElementAt(i - 1)]);
                        roulement.RemoveAt(i - 2);
                    }
                }
            }
        }
    }
}
