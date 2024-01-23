namespace Csharp_eml
{
	internal class Program
	{
		static void Main(string[] args)
		{
			Console.OutputEncoding = System.Text.Encoding.UTF8;
			Console.Title = "Test_eml";
			Console.WriteLine("Début du programme\n");


			(new ExplorerEml()).Demarrer();


			Console.WriteLine("\nFin de programme");
			Console.Read();
			Environment.Exit(0);
		}
	}
}
