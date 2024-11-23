using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using L3S.FromExcelToListClass;
using L3S.FromExcelToListClass.Models;

namespace L3S.TestConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //var summary = BenchmarkRunner.Run<MiBenchmarckAnashei>();
            //Console.WriteLine(summary);
            MiBenchmarckAnashei.MiFuncion();
        }
    }

    public class MiBenchmarckAnashei
    {
        [Benchmark]
        public static void MiFuncion()
        {
            var path = Path.GetFullPath("");
            var directory = new DirectoryInfo(path);

            foreach (var file in directory.GetFiles())
            {
                var resultado = new FromExcelToListClass<PublicacionExcelDTO>().ParseExcelToClass(file);

                Console.WriteLine(resultado.ErrorMessage);
            }
        }
    }
}
