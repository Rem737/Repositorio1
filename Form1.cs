using Microsoft.Office.Interop.Word;
using MicrosoftWord = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.CodeDom.Compiler;
using System.IO;
using Microsoft.CSharp;

namespace WordAsIDE
{
    
    public partial class Form1 : Form
    {
        private string text;
        private Document wordDoc;
        private MicrosoftWord.Application wordApp;
        private string fileName;
        private string executableName;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            wordApp = new MicrosoftWord.Application();
            wordApp.Visible = true;


            wordDoc = wordApp.Documents.Add();
            wordDoc.SaveAs(@"C:\Users\gianl\Desktop\WordAsIDE\MiDocumento.docx");
            wordDoc.Content.Select();
            text = wordDoc.Content.Text;

            wordApp.WindowSelectionChange += new ApplicationEvents4_WindowSelectionChangeEventHandler(WindowSelectionChange);

        }

        private void WindowSelectionChange(Selection sel)
        {
            if (wordDoc != null)
            {
                text = wordDoc.Content.Text;

                MicrosoftWord.Range range = wordDoc.Range();

                Find findObj = wordDoc.Content.Find;
                findObj.ClearFormatting();
                findObj.Text = "//*^p";
                findObj.Forward = true;
                findObj.Wrap = WdFindWrap.wdFindContinue;
                findObj.Format = true;
                findObj.Font.Color = WdColor.wdColorGray25;
                findObj.Execute();

                // Buscar y cambiar el color de las líneas que comienzan por "#"
                findObj.ClearFormatting();
                findObj.Text = "#*^p";
                findObj.Forward = true;
                findObj.Wrap = WdFindWrap.wdFindContinue;
                findObj.Format = true;
                findObj.Font.Color = WdColor.wdColorGreen;
                findObj.Execute();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            if (wordDoc != null)
            {
                wordDoc.Close();
            }
            if (wordApp != null)
            {
                wordApp.Quit();
            }
        }

        private void compileTextOnClick(object sender, EventArgs e)
        {
            string exampleCode = "#include <iostream>\n\nusing namespace std;\n\nint main()\n{\n   cout << \"Hello, World!\" << endl;\n   return 0;\n}\n";
            fileName = @"C:\Users\gianl\Desktop\WordAsIDE\" + "codigo.cpp";
            executableName = @"C:\Users\gianl\Desktop\WordAsIDE\" + "codigo.exe";

            string mingwPath = "C:\\MinGW\\bin";
            Environment.SetEnvironmentVariable("PATH", Environment.GetEnvironmentVariable("PATH") + ";" + mingwPath);


            using (StreamWriter archivo = new StreamWriter(fileName))
            {
                archivo.Write(text);
            }

            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = false;
            process.Start();

            process.StandardInput.WriteLine($"g++ {fileName} -o {executableName}");
            process.StandardInput.Flush();
            process.StandardInput.Close();
            process.WaitForExit();

        }

        private void ExecuteButtonOnClick(object sender, EventArgs e)
        {
            Process process = new Process();
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.CreateNoWindow = false;
            process.Start();

            process.StandardInput.WriteLine($"{executableName}");
            process.StandardInput.Flush();
            process.StandardInput.Close();
            process.WaitForExit();

        }

        private void CompileExecuteOnCLick(object sender, EventArgs e)
        {
            compileTextOnClick(sender, e);
            ExecuteButtonOnClick(sender, e);
        }
    }
}