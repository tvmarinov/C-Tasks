using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Factorials
{
    class Factorials
    {
        static void Main(string[] args)
        {
            int i,j;
            long fact = 1;

            var watch = System.Diagnostics.Stopwatch.StartNew();

            watch.Start();

            for (i = 0; i <= 1000; i++)
            {
                for(j = 0; j <= 10; j++)
                {
                    fact = Factorials.CalculateFactorialRec(j);
                    Console.Write(fact + " ");
                }
                Console.WriteLine();
            }

            watch.Stop();

            Console.WriteLine();
            Console.WriteLine(watch.ElapsedMilliseconds);
            
        }

        //public static int CalculateFactorial(int factorial)
        //{
        //    int i,result = 1;
        //    for (i = 0; i <= factorial; i++)
        //    {
        //        if (i > 1)
        //        {
        //            result *= i;
        //        }
        //    }

        //    return result;
        //}

        public static int CalculateFactorialRec(int factorial)
        {
            if (factorial <= 1)
            {
                return 1;
            }
            return factorial * Factorials.CalculateFactorialRec(factorial - 1);
        }
    }
}
