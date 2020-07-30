using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace automatedSystem
{
    class MyCustomValidation
    {
      
            public static bool ValidLength(string txt, int min, int max) //Checks length & if empty
            {
                bool ok = true;

                if (string.IsNullOrEmpty(txt))
                    ok = false;
                else if (txt.Length < min || txt.Length > max)
                    ok = false;

                return ok;
            }

            public static bool ValidNumber(string txt) //Checks if Number
            {
                bool ok = true;

                for (int x = 0; x < txt.Length; x++)
                {
                    if (!(char.IsNumber(txt[x])))
                    {
                        ok = false;
                    }
                }

                return ok;
            }

            public static bool ValidLetter(string txt) //Checks if input contains Alphabetic Characters 
            {
                bool ok = true;

                if (txt.Trim().Length == 0)
                {
                    ok = false;
                }
                else
                {
                    for (int x = 0; x < txt.Length; x++)
                    {
                        if (!(char.IsLetter(txt[x])))
                        {
                            ok = false;
                        }
                    }
                }
                return ok;
            }
            public static bool ValidLetterNumberWhitespace(string txt) //Allows for Whitespace and Alphabetic Characters
            {
                bool ok = true;

                if (txt.Trim().Length == 0)
                {
                    ok = false;
                }
                else
                {
                    for (int x = 0; x < txt.Length; x++)
                    {
                        if (!(char.IsLetter(txt[x])) && !(char.IsWhiteSpace(txt[x])) && !(char.IsNumber(txt[x])))
                        {
                            ok = false;
                        }
                    }
                }
                return ok;
            }

            public static bool ValidLetterWhitespace(string txt) //Allows for Whitespace and Alphabetic Characters
            {
                bool ok = true;

                if (txt.Trim().Length == 0)
                {
                    ok = false;
                }
                else
                {
                    for (int x = 0; x < txt.Length; x++)
                    {
                        if (!(char.IsLetter(txt[x])) && !(char.IsWhiteSpace(txt[x])))
                        {
                            ok = false;
                        }
                    }
                }
                return ok;
            }


            public static bool ValidForname(string txt) //Allows Alphabetic Characters, dash and whitespace
            {
                bool ok = true;
                if (txt.Trim().Length == 0)
                {
                    ok = false;
                }
                else
                {
                    for (int x = 0; x < txt.Length; x++)
                    {
                        if (!(char.IsLetter(txt[x])) && !(char.IsWhiteSpace(txt[x])) && !(txt[x].Equals('-')))
                        {
                            ok = false;
                        }
                    }
                }
                return ok;
            }

            public static bool ValidSurname(string txt) //Allows Alphabetic Characters, dash, Characters & whitespace
            {
                bool ok = true;
                if (txt.Trim().Length == 0)
                {
                    ok = false;
                }
                else
                {
                    for (int x = 0; x < txt.Length; x++)
                    {
                        if (!(char.IsLetter(txt[x])) && !(char.IsWhiteSpace(txt[x])) && !(txt[x].Equals('-')) && !(txt[x].Equals('\'')))
                        {
                            ok = false;
                        }
                    }
                }
                return ok;
            }


            public static bool ValidEmail(string txt) //Allows Whitespace & Alphanumeric
            {
                bool ok = true;

                if (txt.Trim().Length == 0)
                {
                    ok = false;
                }
                else
                {
                    for (int x = 0; x < txt.Length; x++)
                    {
                        if (!(char.IsLetter(txt[x])) && !(char.IsNumber(txt[x])) && !(txt[x].Equals('@')) && !(txt[x].Equals('-'))
                            && !(txt[x].Equals('_')) && !(txt[x].Equals('.')))
                        {
                            ok = false;
                        }
                    }
                }

                return ok;
            }

            public static bool ValidDOB(DateTime dogDOB) //Checks if Dogs DOB isnt <8 weeks
            {
                bool ok = true;

                DateTime currentDate = DateTime.Now;
                TimeSpan t = currentDate - dogDOB;
                double NoOfDays = t.TotalDays;
                if (NoOfDays <= 56)
                {
                    ok = false;
                }

                return ok;
            }

            public static String FirstLetterEachWordToUpper(String word) // Changes first letter in a word to Uppercase
            {
                Char[] array = word.ToCharArray();

                if (Char.IsLower(array[0]))
                {
                    array[0] = Char.ToUpper(array[0]);
                }

                for (int x = 1; x < array.Length; x++)
                {
                    if (array[x - 1] == ' ')
                    {
                        if (Char.IsLower(array[x]))
                        {
                            array[x] = Char.ToUpper(array[x]);
                        }
                    }
                    else
                    {
                        array[x] = Char.ToLower(array[x]);
                    }
                }
                return new String(array);
            }

            public static String EachLetterToUpper(String word) //Changes all Letters to UpperCase
            {
                Char[] aray = word.ToCharArray();

                for (int x = 0; x < aray.Length; x++)
                {
                    if (Char.IsLower(aray[x]))
                    {
                        aray[x] = Char.ToUpper(aray[x]);
                    }
                }
                return new string(aray);
            }

        public static String RemoveWhiteSpace(String word)
        {
            Char[] array = word.ToCharArray();
            string poly = "";
            for (int i = 0; i < array.Length; i++)
            {
                if (array[i] != ' ')
                {
                    poly += array[i];
                }
            }
            return poly;
        }

        public static bool ValidFileName(string txt)
        {
            bool ok = true;

            if (txt.Trim().Length == 0)
            {
                ok = false;
            }
            else
            {
                for (int x = 0; x < txt.Length; x++)
                {
                    if (!(char.IsLetter(txt[x])) && !(char.IsWhiteSpace(txt[x])) && !(char.IsNumber(txt[x])))
                    {
                        ok = false;
                    }
                }
            }
            return ok;

        }

        
    }
}
