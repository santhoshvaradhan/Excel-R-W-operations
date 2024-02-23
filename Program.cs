using System;
using System.Linq;
using IronXL;
using System.Collections.Generic;
using System.CodeDom.Compiler;
namespace ConsoleApp1
{
     public class Program
        
    {
      
        public  static List<string> firstlist = new List<string>();
        public  static List<string> secondlist = new List<string>();
        public static string messagebody = "";
        public static string mobilenumber = "";
        public static string Decisionfunc(string mark)
        {
            if (mark == "a" || mark == "A" || mark == "absent" || mark == "Absent")
            {
                return mark + "(Fail)";

            }
            else
            {

                if (Convert.ToInt32(mark) >= 50)
                {
                    return mark + "(Pass)";
                }
                else
                {
                    return mark + "(Fail)";
                }
               
            }

        }
        public static void Messageframe(List<string> keyheader,List<string> vlaues)
        {
            for (int i = 0; i < keyheader.Count; i++)
            {
                if (keyheader[i]== "SECTION" || keyheader[i] == "ENROLLNO" || keyheader[i] == "NAME" )
                {
                    messagebody += String.Format("{0}--{1}\n", keyheader[i], vlaues[i]);
                }
                else if(keyheader[i] == "MOBILENUMBER")
                 {
                    mobilenumber =vlaues[i];
          
                }
                else
                {

                    messagebody+=String.Format("{0}--{1}\n",keyheader[i],Decisionfunc(vlaues[i]));
                }

            }
            Console.WriteLine("\nHI THIS IS MESSAGE FROM MAILAM ENGINEERING COLLEGE\nIAT-1 RESULT\n"+messagebody+"Thank you!");
            Console.WriteLine(mobilenumber);
            messagebody = string.Empty;
            mobilenumber = string.Empty;



        }
        

        static void Main()
        {
            WorkBook workbook = WorkBook.Load(filename: "C:\\Users\\Santh\\Downloads\\2023-II-CSE.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();
            var row1 = sheet[Convert.ToString(sheet.GetRow(0).RangeAddress)];
            
            foreach (var value in row1)
            {
                firstlist.Add(value.ToString());

            }

            foreach (var eachRow in sheet.Rows)
            {
                
                if (Convert.ToString(sheet.GetRow(0).RangeAddress) != Convert.ToString(eachRow.RangeAddress))
                {
                    //Console.WriteLine(eachRow);
                    foreach (var cell in eachRow)
                    {
                        //Console.WriteLine(cell.Value);
                      
                        secondlist.Add(cell.ToString());
                    }
                   // Console.WriteLine("sending...");
                    Messageframe(firstlist ,secondlist);
                    secondlist.Clear();
                    //Console.WriteLine("sended...");


                }
                
            }
            Console.ReadKey();
        }
    }
}
  