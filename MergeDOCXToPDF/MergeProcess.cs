using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Linq;

namespace MergeDOCXToPDF
{
	internal class MergeProcess
	{
		private string outWord;
		private string outPDF;

		public MergeProcess(string[] files)
		{
			if (CheckFiles(files))
			{
				outWord = Directory.GetCurrentDirectory() + "\\mergedWord" + (Directory.GetFiles(Directory.GetCurrentDirectory()).Count(a => a.Contains("mergedWord")) + 1) + ".docx";
				outPDF = Directory.GetCurrentDirectory() + "\\" + GetPDFOutputName() + ".pdf";

				Console.WriteLine("Please give every file a numerical ID to sort them:");
				Dictionary<int, string> sortedFiles = GetFileOrder(files);
				Console.WriteLine("Merging files...");
				MergeDOCX(sortedFiles);
				Console.WriteLine("Converting to PDF...");
				ConvertDOCXToPDF(outWord);
				Console.WriteLine("Deleting temporary files...");
				File.Delete(outWord);

				Console.WriteLine($"Finished! Output file is \"{outPDF}\"");
			}

			Console.WriteLine();
			Console.Write("Closing window in... 10");
			System.Threading.Thread.Sleep(1000);
			Console.CursorLeft -= 2;
			Console.Write("  ");
			Console.CursorLeft--;
			for (int i = 9; i >= 0; i--)
			{
				Console.CursorLeft--;
				Console.Write(i);
				System.Threading.Thread.Sleep(1000);
			}
		}

		private Dictionary<int, string> GetFileOrder(string[] files)
		{
			var dict = new Dictionary<int, string>();
			for (int i = 0; i < files.Length; i++)
			{
				int id;
				while (true)
				{
					Console.Write($"\"{files[i]}\": ");

					char pressedKey = Console.ReadKey().KeyChar;
					Console.WriteLine();

					if (int.TryParse(pressedKey.ToString(), out id))
					{
						break;
					}

					Console.WriteLine($"{pressedKey} is not a number. Please try again.");
				}

				dict.Add(id, files[i]);
			}
			return dict;
		}

		private string GetPDFOutputName()
		{
			Console.Write("Name of PDF file: ");
			return Console.ReadLine();
		}

		private bool CheckFiles(string[] args)
		{
			for (int i = 0; i < args.Length; i++)
			{
				if (!File.Exists(args[i]))
				{
					Console.WriteLine($"File \"{args[i]}\" does not exist. Canceled process.");
					return false;
				}

				if (!args[i].EndsWith(".docx"))
				{
					Console.WriteLine($"File \"{args[i]}\" is not a word file. Canceled process.");
					return false;
				}
			}

			if (args.Length <= 0)
			{
				Console.WriteLine("No arguments specified. Canceled process.\n");
				Console.WriteLine("How to use this program:\n" +
				                  "This program converts DOCX files into a singular PDF file.\n" +
				                  "To do this, either give the program one argument for each\n" +
				                  "file or simply drag&drop them into this application\n(the EXE, not the command window).");
				return false;
			}

			return true;
		}

		private void MergeDOCX(Dictionary<int, string> filesWithID)
		{
			var app = new Word.Application();
			Word.Document doc = app.Documents.Add();
			Word.Selection sel = app.Selection;

			int[] order = filesWithID.Select(a => a.Key).OrderBy(b => b).ToArray();

			for (int i = 0; i < order.Length; i++)
			{
				sel.InsertFile(filesWithID[order[i]]);
				if (i < filesWithID.Count - 1)
				{
					sel.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
				}
			}

			doc.SaveAs2(outWord);

			GC.Collect();
			GC.WaitForPendingFinalizers();

			Marshal.ReleaseComObject(sel);
			doc.Close();
			Marshal.ReleaseComObject(doc);
			app.Quit();
			Marshal.ReleaseComObject(app);
		}

		private void ConvertDOCXToPDF(string docxFile)
		{
			var app = new Word.Application();
			Word.Document doc = app.Documents.Open(docxFile);

			doc.ExportAsFixedFormat(outPDF, Word.WdExportFormat.wdExportFormatPDF);

			GC.Collect();
			GC.WaitForPendingFinalizers();
			
			doc.Close();
			Marshal.ReleaseComObject(doc);
			app.Quit();
			Marshal.ReleaseComObject(app);
		}
	}
}
