using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using  System.Reflection;

namespace Execute_Selected_Tests
{
    public partial class Form1 : Form
    {
        private bool processesAreRunning =true;
        public string nameOfAutomatedCase = string.Empty;
        public bool startStopConditon = true;
        public Process process;
        private int testsRun;
        private int passed;
        public int progress;
        public int errorCount;
        private bool testPassedFlag;
        public int failureCount;
        private int notRunCount;
        public string currentDir = Directory.GetParent(Directory.GetParent(Environment.CurrentDirectory).FullName).FullName;//@"D:\SVN\ComTests\Develop\ComTests";
        public string optionsFile = Path.Combine(Directory
            .GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)).FullName,"Local","Execute Selected Tests", "Options.txt");
        public string buildFileName = string.Empty;
        public string tempContent;
        public string fileToOpen = string.Empty;
        public List<Tuple<string, string>> list;
        public List<string> listFailures;

        public string appData = Directory
            .GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)).FullName;
        public Form1()
        {
            InitializeComponent();
        }

        private List<Tuple<string, string>> GetFailedTestsList(string[] content)
        {
            var list01 = new List<Tuple<string, string>>();
            for (var line = 0; line < content.Length; line++)
            {
                if (content[line].Contains(")"))
                {
                    if (content[line].Contains("Failure :") || content[line].Contains("Error :"))
                    {
                        var testName = content[line].Substring(content[line].IndexOf(":", StringComparison.Ordinal) + 2);
                        var dllName = testName.Remove(testName.LastIndexOf(".", StringComparison.Ordinal));
                        dllName = dllName.Remove(dllName.LastIndexOf(".", StringComparison.Ordinal) + 1) + "dll";
                        list01.Add(Tuple.Create(dllName, testName));
                    }
                }
            }

            return list01;
        }

        private string[] CleanData(string[] content)
        {
            for (var line = 0; line < content.Length; line++)
                if (!content[line].Contains("Errors and Failures:"))
                {
                    content[line] = string.Empty;
                }
                else
                {
                    content[line] = string.Empty;
                    content = content.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                    break;
                }

            return content;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox3.Items.Add("nunit-console-x64.exe");
            comboBox3.Items.Add("nunit-console.exe");
            comboBox3.SelectedItem = "nunit-console.exe";

            if (File.Exists(optionsFile).Equals(true))
            {
                var pathSet = false;
                var exeNunit = false;
                var content = File.ReadAllLines(optionsFile);
                if (content.Length > 0)
                {
                    for (var index = 0; index < content.Length; index++)
                    {
                        if (content[index].Contains("PATH"))
                        {
                            textBox3.Text=content[index].Substring(content[index].IndexOf("=")+1);
                            pathSet = true;
                        }
                        
                        if (content[index].Contains("NUNITEXE"))
                        {
                            comboBox3.SelectedItem = content[index].Substring(content[index].IndexOf("=") + 1);
                            exeNunit = true;
                        }
                        
                    }
                }
                if(exeNunit==false)
                {
                    comboBox3.SelectedItem = "nunit-console.exe";
                }
                if (pathSet == false)
                {
                    textBox3.Text = @"C:\Program Files (x86)\NUnit 2.6.4\bin";
                }
            }
            else
            {
                textBox3.Text = @"C:\Program Files (x86)\NUnit 2.6.4\bin";
            }

            button4.Enabled = false;
            button3.Enabled = false;
            button2.Enabled = false;
            comboBox1.Enabled = false;
            comboBox2.Enabled = false;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox1.Items.Add("x64");
            comboBox1.Items.Add("x86");
            comboBox1.SelectedItem = "x64";
            comboBox2.Items.Add("Debug");
            comboBox2.Items.Add("DebugRt");
            comboBox2.Items.Add("CodeCoverage");
            comboBox2.Items.Add("CodeCoverageRt");
            comboBox2.SelectedItem = "Debug";
            progressBar1.Minimum = 0;
            progressBar1.Value = 0;
            progressBar1.Style = ProgressBarStyle.Continuous;
            textBox1.Enabled = false;
            textBox2.Enabled = true;
            textBox1.Text = "7200000";
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private Point lastPoint;

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.Left += e.X - lastPoint.X;
                this.Top += e.Y - lastPoint.Y;
            }
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            lastPoint = new Point(e.X, e.Y);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button3_Click(sender, e);
            Application.Exit();
        }

        private BackgroundWorker worker;
        private BackgroundWorker worker04;
        private void AsyncHelper(int count)
        {
            var slnPath = Path.Combine(currentDir, "CommonTests.sln");
            if (File.Exists(slnPath).Equals(true))
            {
                errorCount = 0;
                testsRun = 0;
                failureCount = 0;
                notRunCount = 0;
                passed = 0;
                if (fileToOpen.Equals(string.Empty) && label14.Text.Equals(string.Empty))
                {
                    MessageBox.Show("Select a file Or Type a Test Case Name before starting the execution.");
                }
                else
                {
                    Invoke(new Action(() => { pictureBox2.Enabled = false; }));
                    Invoke(new Action(() => { button5.Enabled = false; }));
                    Invoke(new Action(() => { textBox3.Enabled = false; }));
                    Invoke(new Action(() => { button4.Enabled = false; }));
                    Invoke(new Action(() => { button3.Enabled = true; }));
                    Invoke(new Action(() => { button2.Enabled = false; }));
                    Invoke(new Action(() => { comboBox1.Enabled = false; }));
                    Invoke(new Action(() => { comboBox2.Enabled = false; }));
                    Invoke(new Action(() => { textBox1.Enabled = false; }));
                    Invoke(new Action(() => { textBox2.Enabled = false; }));
                    Invoke(new Action(() => { button1.Enabled = false; }));
                    Invoke(new Action(() => { button2.Enabled = false; }));
                    Invoke(new Action(() => { button4.Enabled = false; }));
                    progress = 0;
                    Invoke(new Action(() => { label9.Text = $"Tests Run {testsRun}/{list.Count}"; }));
                    Invoke(new Action(() =>
                    {
                        label12.Text = $"Not Run:{notRunCount}, Passed:{progress}, Errors:{errorCount}, Failures:{failureCount}";
                    }));
                    Invoke(new Action(() => { progressBar1.Value = 0; }));
                    string arc = string.Empty;
                    Invoke(new Action(() => { arc = comboBox1.SelectedItem.ToString(); }));
                    Invoke(new Action(() => { comboBox1.Enabled = false; }));
                    string config = string.Empty;
                    Invoke(new Action(() => { config = comboBox2.SelectedItem.ToString(); }));

                    string tempFile = string.Empty;
                    foreach (var value in list)
                    {
                        testPassedFlag = false;
                        var dllPath = $@"{currentDir}\bin\{arc}\{config}\{value.Item1}";
                        if (File.Exists(dllPath) == false)
                        {
                            MessageBox.Show(
                                $"{value.Item1} was not found.\nPlease build the solution again before executing.");

                            break;
                        }

                        var timeOut = Convert.ToInt32(textBox1.Text);
                        var psContent = $"cd \"{textBox3.Text}\";.\\{comboBox3.SelectedItem.ToString()} /run=\"{value.Item2}\" /timeout={timeOut} \"{dllPath}\"";

                        var psFile = Path.Combine(Environment.CurrentDirectory, "Temp", "Script.ps1");
                        DeleteFileIfExists(psFile);

                        File.WriteAllText(psFile, psContent);
                        process = new Process();
                        process.StartInfo = new ProcessStartInfo()
                        {
                            FileName = "powershell.exe",
                            Arguments = $"-NoProfile -ExecutionPolicy unrestricted \"{psFile}\"",
                            UseShellExecute = false,
                            WindowStyle = ProcessWindowStyle.Hidden,
                            RedirectStandardOutput = true,
                            CreateNoWindow = true
                        };
                        process.Start();
                        var output = process.StandardOutput.ReadToEnd();
                        tempFile = Path.Combine(Environment.CurrentDirectory, "Temp", "TempFile.txt");
                        DeleteFileIfExists(tempFile);
                        File.WriteAllText(tempFile, output);
                        var content = File.ReadAllLines(tempFile);
                        for (var index = 0; index < content.Length; index++)
                        {
                            if (content[index].Contains("Tests run: 0,"))
                            {
                                notRunCount++;
                                content[index] = "Errors and Failures:";
                                content[index+1]= $" ) Test Not Run : {value.Item2}" + Environment.NewLine + "    Test was not run, maybe the Test Name was invalid or Configuration selected for the test was wrong.";
                                testPassedFlag = true;
                            }
                        }

                        content = CleanData(content);
                        var colonCounter = 0;
                        for (var index = 0; index < content.Length; index++)
                        {
                            if (content[index].Contains("Failure :"))
                            {
                                testPassedFlag = true;
                                if (colonCounter == 0)
                                {
                                    failureCount++;
                                    testsRun++;
                                    if (content[0].Contains("Test Failure :").Equals(false))
                                    {
                                        content[0] = $"1) Test Failure : {value.Item2}" + Environment.NewLine;
                                    }
                                }
                                if (colonCounter > 0)
                                {
                                    content[index] = content[index].Replace(":", string.Empty);
                                }
                                colonCounter++;
                            }
                            if(content[index].Contains("Error :"))
                            {
                                testPassedFlag = true;
                                if (colonCounter == 0)
                                {
                                    testsRun++;
                                    errorCount++;
                                    if (content[0].Contains("Test Error :").Equals(false))
                                    {
                                        content[0] = $"1) Test Error : {value.Item2}" + Environment.NewLine;
                                    }
                                }
                                if (colonCounter>0)
                                {
                                    content[index] = content[index].Replace(":", string.Empty);
                                }
                                colonCounter++;
                            }
                        }

                        File.WriteAllLines(tempFile, content);
                        var contentFinal = File.ReadAllText(tempFile);
                        tempContent = tempContent + "\n" + contentFinal;
                        DeleteFileIfExists(tempFile);
                        DeleteFileIfExists(psFile);
                        if (startStopConditon.Equals(true))
                        {
                            if (testPassedFlag.Equals(false))
                            {
                                passed++;
                            }
                            progress++;
                            Invoke(new Action(() => { progressBar1.Value = progress; }));
                            Invoke(new Action(() => { label9.Text = $"Tests Run {passed+failureCount+errorCount}/{list.Count}"; }));
                            Invoke(new Action(() =>
                            {
                                label12.Text = $"Not Run :{notRunCount},Tests Run:{passed+errorCount+failureCount},Passed:{passed}, Errors:{errorCount}, Failures:{failureCount}\n"; ;
                            }));

                        }

                        if (startStopConditon.Equals(false))
                        {
                            worker04 = new BackgroundWorker();
                            worker04.DoWork += (obj, ea) => AsyncHelperExit(1);
                            worker04.RunWorkerAsync();
                            break;
                        }
                    }

                    var finalContent = TesCaseNumbering(tempContent);
                    finalContent = RemoveEmptyLines(finalContent);
                    var outputFile = Path.Combine(Environment.CurrentDirectory, "TestResults.txt");
                    File.WriteAllLines(outputFile, finalContent);
                    var processStartInfo = new ProcessStartInfo()
                    {
                        UseShellExecute = true,
                        FileName = "notepad.exe"
                    };
                    Process.Start("notepad.exe", outputFile);
                    DeleteFileIfExists(tempFile);
                    Invoke(new Action(() => { button2.Enabled = true; }));
                    Invoke(new Action(() => { button4.Enabled = true; }));
                    Invoke(new Action(() => { pictureBox2.Enabled = true; }));
                    Invoke(new Action(() => { comboBox1.Enabled = true; }));
                    Invoke(new Action(() => { comboBox2.Enabled = true; }));
                    Invoke(new Action(() => { textBox1.Enabled = true; }));
                    Invoke(new Action(() => { button1.Enabled = true; }));
                    Invoke(new Action(() => { button2.Enabled = true; }));
                    Invoke(new Action(() => { button3.Enabled = false; }));
                    Invoke(new Action(() => { button4.Enabled = true; }));
                    if (textBox2.Text != string.Empty)
                    {
                        Invoke(new Action(() => { button1.Enabled = false; }));
                        Invoke(new Action(() => { textBox2.Enabled = true; }));
                    }
                    else
                    {
                        Invoke(new Action(() => { button1.Enabled = true; }));
                        Invoke(new Action(() => { textBox2.Enabled = false; }));
                    }
                    Invoke(new Action(() => { button5.Enabled = true; }));
                    Invoke(new Action(() => { textBox3.Enabled = true; }));
                }
            }
            else
            {
                MessageBox.Show("Move the Exe File to ComTests Directory to start execution.");
            }
            
        }

        private string[] TesCaseNumbering(string data)
        {
            if (textBox2.Text != string.Empty && notRunCount == 1)
            {
                MessageBox.Show("Test was not run. Test name was invalid Or Selected Configuration was wrong.");
            }
            if (textBox2.Text == string.Empty && notRunCount >0)
            {
                MessageBox.Show($"{notRunCount} Test(s) did not run. Test name was invalid Or Selected Configuration was wrong for the test.");
            }

            var header = $"Errors and Failures:    →  Not Run :{notRunCount},Tests Run:{passed + errorCount + failureCount},Passed:{passed}, Errors:{errorCount}, Failures:{failureCount}\n";
            data = header + data;
            var tempFile = Path.Combine(Environment.CurrentDirectory, "Temp", "TempFile.txt");
            DeleteFileIfExists(tempFile);
            File.WriteAllText(tempFile, data);
            var content = File.ReadAllLines(tempFile);
            var counter = 1;
            passed = 0;
            errorCount = 0;
            testsRun = 0;
            failureCount = 0;
            notRunCount = 0;
            for (var index = 0; index < content.Length; index++)
            {
                if (content[index].Contains("Failure :"))
                {
                    content[index] = $"{counter}){content[index].Substring(content[index].IndexOf(")") + 1)}";
                    counter++;
                    failureCount++;
                }

                if (content[index].Contains("Error :"))
                {
                    content[index] = $"{counter}){content[index].Substring(content[index].IndexOf(")") + 1)}";
                    counter++;
                    errorCount++;
                }

                if (content[index].Contains(") Test Not Run :"))
                {
                    content[index] = $"{counter}){content[index].Substring(content[index].IndexOf(")")+1)}";
                    counter++;
                    notRunCount++;
                }
            }

            
            return content;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            tempContent = string.Empty;
            worker = new BackgroundWorker();
            worker.DoWork += (obj, ea) => AsyncHelper(1);
            worker.RunWorkerAsync();
        }

        private void DeleteFileIfExists(string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        private void BrowseHelper(string fileToOpen)
        {
            var listFail01 = new List<string>();
            bool extensionCondition = false;
            if (!Path.GetExtension(fileToOpen).Equals(".txt") && !Path.GetExtension(fileToOpen).Equals(".xml"))
            {
                extensionCondition = true;
            }

            var list01 = new List<Tuple<string, string>>();
            if (extensionCondition.Equals(true))
            {
                MessageBox.Show("Please select a Text File(.txt) Or Xml File(.xml)");
            }
            else
            {
                var content01 = File.ReadAllLines(fileToOpen);
                content01 = CleanData(content01);
                if (Path.GetExtension(fileToOpen).Equals(".txt"))
                {
                    list = GetFailedTestsList(content01);
                }
                else
                {
                    var xmlDoc = new XmlDocument();
                    xmlDoc.Load(fileToOpen);
                    
                    var testCaseNodes = xmlDoc.GetElementsByTagName("test-case");
                    foreach (XmlNode value in testCaseNodes)
                    {
                        if (value.SelectSingleNode("failure") != null)
                        {
                            listFail01.Add(value.SelectSingleNode("failure").InnerText);
                            var testName = value.Attributes[0].Value;
                            
                            var dllName = testName.Remove(testName.LastIndexOf(".", StringComparison.Ordinal));
                            dllName = dllName.Remove(dllName.LastIndexOf(".", StringComparison.Ordinal) + 1) +
                                      "dll";
                            list01.Add(Tuple.Create(dllName,testName));
                        }
                    }
                }

                if (Path.GetExtension(fileToOpen).Equals(".xml"))
                {
                    list = list01;
                    listFailures = listFail01;
                }

                if (list.Count > 0)
                {
                    Invoke(new Action(() => { label10.Text = $"Selected File : {fileToOpen}"; }));
                    Invoke(new Action(() => { label3.Text = $"{list.Count} Test(s) will be executed."; }));
                    Invoke(new Action(() => { comboBox1.Enabled = true; }));
                    Invoke(new Action(() => { comboBox2.Enabled = true; }));
                    Invoke(new Action(() => { button2.Enabled = true; }));
                    Invoke(new Action(() => { button4.Enabled = true; }));
                    Invoke(new Action(() => { textBox1.Enabled = true; }));
                    Invoke(new Action(() => { textBox2.Enabled = true; }));
                    Invoke(new Action(() => { progressBar1.Maximum = list.Count; }));
                    Invoke(new Action(() => { label9.Text = string.Empty; }));
                    Invoke(new Action(() => { label12.Text = string.Empty; }));
                    Invoke(new Action(() => { textBox2.Enabled = false; }));
                }
                else
                {
                    Invoke(new Action(() => { textBox2.Enabled = true; }));
                    Invoke(new Action(() => { label10.Text = "No File Selected"; }));
                    Invoke(new Action(() => { label3.Text = "0 Test(s) will be executed."; }));
                    Invoke(new Action(() => { button2.Enabled = false; }));
                    Invoke(new Action(() => { button4.Enabled = false; }));
                    Invoke(new Action(() => { comboBox1.Enabled = false; }));
                    Invoke(new Action(() => { comboBox2.Enabled = false; }));
                    MessageBox.Show("No tests were found in the selected file\n\nSelect a different text file.");
                }
            }
        }

        private BackgroundWorker openFile;
        private void button1_Click_2(object sender, EventArgs e)
        {
            tempContent = string.Empty;
            progressBar1.Value = 0;
            var FD = new OpenFileDialog();
            FD.Filter = "Text Files (*.txt)|*.txt|Xml files (*.*)|*.xml";
            FD.FilterIndex = 2;
            FD.RestoreDirectory = true;
            if (FD.ShowDialog() == DialogResult.OK)
            {
                fileToOpen = FD.FileName;
                openFile = new BackgroundWorker();
                openFile.DoWork += (obj, ea) => BrowseHelper(fileToOpen);
                openFile.RunWorkerAsync();
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {
        }

        private void button55_MouseHover(object sender, EventArgs e)
        {

        }

        private BackgroundWorker worker02;

        private void AsyncHelperExit(int count)
        {
            var psFile = Path.Combine(Environment.CurrentDirectory, "Temp", "TaskKiller.ps1");
            DeleteFileIfExists(psFile);
            for (var counter = 1; counter <=5; counter++)
            {
                var tempProcess = Process.GetProcessesByName("powershell.exe");
                    if (tempProcess.Length > 0)
                    {
                        foreach (var value in tempProcess)
                        {
                            value.Kill();
                        }
                    }
                    var psContent = "Stop-Process -Name \"IPEmotion\" -Force;Stop-Process -Name \"IPEmotionRT.UI\" -Force;Stop-Process -Name \"IPEmotionSettings\" -Force;Stop-Process -Name \"IPEmotionSettingsRT.UI\" -Force;Stop-Process -Name \"nunit\" -Force;Stop-Process -Name \"nunit-agent\" -Force;Stop-Process -Name \"nunit-console\" -Force;";
                    File.WriteAllText(psFile, psContent);
                    process = new Process();
                    process.StartInfo = new ProcessStartInfo()
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy unrestricted \"{psFile}\"",
                        UseShellExecute = false,
                        WindowStyle = ProcessWindowStyle.Hidden,
                        CreateNoWindow = true
                    };
                    process.Start();
                    Thread.Sleep(200);
                    process.WaitForExit();
            }

            DeleteFileIfExists(psFile);
            startStopConditon = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            startStopConditon = false;
            worker02 = new BackgroundWorker();
            worker02.DoWork += (obj, ea) => AsyncHelperExit(1);
            worker02.RunWorkerAsync();
            Invoke(new Action(() => { button3.Enabled = false; }));
            Invoke(new Action(() => { button4.Enabled = true; }));
            Invoke(new Action(() => { button2.Enabled = true; }));
            Invoke(new Action(() => { comboBox1.Enabled = true; }));
            Invoke(new Action(() => { comboBox2.Enabled = true; }));
            Invoke(new Action(() => { textBox1.Enabled = true; }));
            Invoke(new Action(() => { textBox2.Enabled = true; }));
            Invoke(new Action(() => { textBox3.Enabled = true; }));
            Invoke(new Action(() => { button5.Enabled = true; }));
            if (textBox2.Text != string.Empty)
            {
                button1.Enabled = true;
            }

            if (label10.Text != "No File Selected")
            {
                textBox2.Enabled = true;
            }
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private BackgroundWorker worker03;

        private void BuildSolution()
        {
            var buildProcess = new Process();
            buildProcess.StartInfo.FileName = Path.Combine(currentDir, buildFileName);
            buildProcess.Start();
            buildProcess.WaitForExit();
        }

        private void HelperSelectBuildFile()
        {
            if (comboBox1.SelectedItem == "x86")
            {
                buildFileName = "buildall_x86.cmd";
            }
            else if (comboBox1.SelectedItem == "x64")
            {
                buildFileName = "buildall_x64.cmd";
            }
            

            label11.Text = $"{buildFileName} Selected based on current settings";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            HelperSelectBuildFile();
            worker03 = new BackgroundWorker();
            worker03.DoWork += (obj, ea) => BuildSolution();
            worker03.RunWorkerAsync();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            HelperSelectBuildFile();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            HelperSelectBuildFile();
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click_1(object sender, EventArgs e)
        {
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            var filter = textBox2.Text.Replace(" ", string.Empty);
            textBox2.Text = filter;
            if (textBox2.Text != string.Empty)
            {
                var testName = textBox2.Text;
                var tempVar = string.Empty;
                var character = '.';
                var value = true;
                if (testName.Count(temp => temp.Equals(character)) == 3)
                {
                    var index01 = testName.IndexOf(".");
                    tempVar = testName.Remove(index01, 1);
                    var index02 = tempVar.IndexOf(".");
                    tempVar = tempVar.Remove(index02, 1);
                    var index03 = tempVar.IndexOf(".");
                    if (index03 > index02 + 4 && index02 > index01 + 4)
                    {
                        value = true;
                    }
                    else
                    {
                        value = false;
                    }
                }

                if (testName.Split('.').Length - 1 == 3 && value.Equals(true))
                {
                    var dllName = testName.Remove(testName.LastIndexOf(".", StringComparison.Ordinal));
                    dllName = dllName.Remove(dllName.LastIndexOf(".", StringComparison.Ordinal) + 1) + "dll";
                    label14.Text = "OK";
                    label14.ForeColor = Color.Green;
                    button1.Enabled = false;
                    comboBox1.Enabled = true;
                    comboBox2.Enabled = true;
                    button2.Enabled = true;
                    button4.Enabled = true;
                    textBox1.Enabled = true;
                    textBox2.Enabled = true;
                    progressBar1.Maximum = 1;
                    label9.Text = string.Empty;
                    label12.Text = string.Empty;
                    textBox1.Enabled = true;
                    var list02 = new List<Tuple<string, string>> {Tuple.Create(dllName, testName)};
                    list = list02;
                    if (list.Count > 0)
                    {
                        label10.Text = $"Selected File : {fileToOpen}";
                        label3.Text = $"{list.Count} Test(s) will be executed.";
                        comboBox1.Enabled = true;
                        comboBox2.Enabled = true;
                        button2.Enabled = true;
                        button1.Enabled = false;
                        button4.Enabled = true;
                        textBox1.Enabled = true;
                        textBox2.Enabled = true;
                        progressBar1.Maximum = list.Count;
                        label9.Text = string.Empty;
                        label12.Text = string.Empty;
                        textBox2.Enabled = true;
                        pictureBox2.Enabled = false;
                    }
                    else
                    {
                        pictureBox2.Enabled = true;
                        textBox2.Enabled = false;
                        label10.Text = "No File Selected";
                        label3.Text = "0 Test(s) will be executed.";
                        button2.Enabled = false;
                        button1.Enabled = false;
                        button4.Enabled = false;
                        comboBox1.Enabled = false;
                        comboBox2.Enabled = false;
                        MessageBox.Show("No tests were found in the selected file\n\nSelect a different text file.");
                    }
                }
                else
                {
                    button1.Enabled = false;
                    label14.Text =
                        "Enter a proper name.\nEg: Scripting.Tests.ScriptingTest.IM49578_IdDsSupportConfigurationComOnly";
                    label14.ForeColor = Color.Red;
                    nameOfAutomatedCase = label14.Text;
                }
            }
            else
            {
                pictureBox2.Enabled = true;
                label14.Text = string.Empty;
                button1.Enabled = true;
                label10.Text = "No File Selected";
                label3.Text = "0 Test(s) will be executed.";
                button2.Enabled = false;
                button4.Enabled = false;
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
                textBox1.Enabled = false;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            if (label10.Text != "No File Selected" &textBox2.Text==string.Empty)
            {
                button1.Enabled = true;
                label10.Text = "No File Selected";
                label3.Text = "0 Test(s) will be executed.";
                button2.Enabled = false;
                button4.Enabled = false;
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
                textBox2.Enabled = true;
                label10.Text = "No File Selected";
            }
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void ChangeOptions(string value)
        {
            var dir = Path.Combine(appData, "Local","Execute Selected Tests");
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (!File.Exists(optionsFile))
            {
                File.Create(optionsFile).Close();
            }
            if (File.Exists(optionsFile).Equals(true))
            {
                File.WriteAllText(optionsFile, value);
            }
            
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            var file = Path.Combine(textBox3.Text, comboBox3.SelectedItem.ToString());

            if (File.Exists(file))
            {
                label111.ForeColor=Color.Green;
                label111.Text = "OK";
                button5.Enabled = true;
                button2.Enabled = true;
            }
            else
            {
                label111.ForeColor = Color.Red;
                label111.Text = $"Path does not contain\n {comboBox3.SelectedItem.ToString()}.";
                button5.Enabled = false;
                button2.Enabled = false;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ChangeOptions($"PATH={textBox3.Text}\nNUNITEXE={comboBox3.SelectedItem}");
            label111.Text = "Path saved.";
        }

        public string[] RemoveEmptyLines(string[] content)
        {
            for (var index = 0; index < content.Length; index++)
            {
                if (content[index].Contains("Failure :"))
                {
                    content[index] = "$@" + content[index];
                }

                if (content[index].Contains("Error :"))
                {
                    content[index] = "$@" + content[index];
                }

                if (content[index].Contains(") Test Not Run :"))
                {
                    content[index] = "$@" + content[index];
                }
            }
            var tempFile = Path.Combine(Environment.CurrentDirectory, "Temp", "TempFile.txt");
            DeleteFileIfExists(tempFile);
            File.WriteAllLines(tempFile,content);
            var content03 =File.ReadAllText(tempFile);
            var content02 = content03.Replace("\r\n", string.Empty);
            content02 = content02.Replace("$@", "\r\n\r\n");
            File.WriteAllText(tempFile,content02);
            var content04 = File.ReadAllLines(tempFile);
            foreach (var value in list)
            {
                for (var index = 0; index < content04.Length; index++)
                {
                    if (content04[index].Contains(value.Item2))
                    {
                        content04[index] = content04[index].Replace(value.Item2, value.Item2 + Environment.NewLine);
                    }
                    
                }
            }
            DeleteFileIfExists(tempFile);
            return content04;
        }
        private void label17_Click(object sender, EventArgs e)
        {
            
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// Execute Power shell Script
        /// </summary>
        /// <param name="command"></param>
        public static void ExecuteCommandWithAdminPowerShell(string command)
        {
            var psFile = Path.Combine(Environment.CurrentDirectory, "Temp", "Powershell.ps1");
            if (File.Exists(psFile))
            {
                File.Delete(psFile);
            }
            File.WriteAllText(psFile, command);
            var process = new Process
            {
                StartInfo = new ProcessStartInfo()
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy unrestricted \"{psFile}\"",
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true
                }
            };
            process.Start();
            process.WaitForExit();
        }

        /// <summary>
        /// Kills all processes with a given name
        /// </summary>
        /// <param name="processName">name of the process(es) to kill</param>
        public static void KillAllProcessesByName(string processName)
        {
            var processes = Process.GetProcessesByName(processName);
            foreach (var process in processes)
            {
                process.Kill();
            }
            ExecuteCommandWithAdminPowerShell($"Stop-Process -Name \"{processName}\" -Force");
            Thread.Sleep(1000);
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            failedTestWorker = new BackgroundWorker();
            failedTestWorker.DoWork += (obj, ea) => GetFailureReport();
            failedTestWorker.RunWorkerAsync();
        }

        private BackgroundWorker failedTestWorker;
        private void GetFailureReport()
        {
            if (fileToOpen.Equals(string.Empty))
            {
                MessageBox.Show("Browse a file before clicking Generate.");
            }
            else
            {
                Invoke(new Action(() => { button8.Enabled = false; }));
                Invoke(new Action(() => { button7.Enabled = false; }));
                Invoke(new Action(() => { button6.Enabled = false; }));
                var defaultText = string.Empty;
                Invoke(new Action(() => { defaultText = button6.Text; }));
                Invoke(new Action(() => { button6.Text = "Please Wait.."; }));


                var info = new FileInfo(fileToOpen);
                if (info.Extension.Equals(".xml"))
                {
                    var assembly = Assembly.GetExecutingAssembly();
                    var reader = assembly.GetManifestResourceStream("Execute_Selected_Tests.Files.Template.xlsx");
                    var file = Path.Combine(Environment.CurrentDirectory, "COM Test Update Report.xlsx");
                    KillAllProcessesByName("EXCEL");
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }
                    using (Stream s = File.Create(file))
                    {
                        reader.CopyTo(s);
                    }
                    var xlApplication = new Microsoft.Office.Interop.Excel.Application
                    {
                        Visible = false
                    };

                    var workbook = xlApplication.Workbooks.Open(file);
                    var sheet = workbook.Worksheets[1];
                    for (var count = 0; count < list.Count; count++)
                    {
                        sheet.Cells[count + 2, 1] = count + 1;
                        sheet.Cells[count + 2, 5] = list[count].Item2;
                        sheet.Cells[count + 2, 4] = listFailures[count];
                    }

                    workbook.Save();
                    xlApplication.Visible = true;


                    Invoke(new Action(() => { button6.Enabled = true; }));
                    Invoke(new Action(() => { button6.Text = defaultText; }));
                    Invoke(new Action(() => { button7.Enabled = true; }));
                    Invoke(new Action(() => { button8.Enabled = true; }));
                }

                else
                {
                    MessageBox.Show(
                        "This Feature works with XML File only currently. \nPlease Select XML Results file to generate report");
                }
            }
        }
        

        private void label17_Click_1(object sender, EventArgs e)
        {

        }


        private BackgroundWorker ignoredTestWorker;
        private void button7_Click(object sender, EventArgs e)
        {
            ignoredTestWorker = new BackgroundWorker();
            ignoredTestWorker.DoWork += (obj, ea) => GetIgnoredTests();
            ignoredTestWorker.RunWorkerAsync();
        }

        private void GetIgnoredTests()
        {
            if (fileToOpen.Equals(string.Empty))
            {
                MessageBox.Show("Browse a file before clicking Generate.");
            }
            else
            {
                Invoke(new Action(() => { button7.Enabled = false; }));
                Invoke(new Action(() => { button6.Enabled = false; }));
                Invoke(new Action(() => { button8.Enabled = false; }));
                var defaultText = string.Empty;
                Invoke(new Action(() => { defaultText = button7.Text; }));
                Invoke(new Action(() => { button7.Text = "Please Wait.."; }));
                var info = new FileInfo(fileToOpen);
                if (info.Extension.Equals(".xml"))
                {

                    var xmlDoc = new XmlDocument();
                    xmlDoc.Load(fileToOpen);

                    var testCaseNodes = xmlDoc.GetElementsByTagName("test-case");
                    var tmpTxt = string.Empty;
                    var sentForReview = new List<Tuple<string, string>>();
                    var update = new List<Tuple<string, string>>();
                    var license = new List<Tuple<string, string>>();
                    var freeze = new List<Tuple<string, string>>();
                    var test = new List<Tuple<string, string>>();
                    var incomplete = new List<Tuple<string, string>>();
                    var others = new List<Tuple<string, string>>();
                    foreach (XmlNode node in testCaseNodes)
                    {
                        if (node.Attributes[3].Value.Equals("Ignored"))
                        {
                            var text = node.SelectSingleNode("reason").SelectSingleNode("message").InnerText;
                            if (Regex.IsMatch(text, Regex.Escape("review"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("Bug"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("Discuss"), RegexOptions.IgnoreCase))
                            {
                                sentForReview.Add(Tuple.Create(node.Attributes[0].Value, text));
                            }
                            else if (Regex.IsMatch(text, Regex.Escape("Test Plugin"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("Test-PI"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("Test-PI"), RegexOptions.IgnoreCase))
                            {
                                test.Add(Tuple.Create(node.Attributes[0].Value, text));
                            }
                            else if (Regex.IsMatch(text, Regex.Escape("Incomplete"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("com function"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("Function"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("Feature"), RegexOptions.IgnoreCase))
                            {
                                incomplete.Add(Tuple.Create(node.Attributes[0].Value, text));
                            }
                            else if (Regex.IsMatch(text, Regex.Escape("updat"), RegexOptions.IgnoreCase))
                            {
                                update.Add(Tuple.Create(node.Attributes[0].Value, text));
                            }
                            else if (Regex.IsMatch(text, Regex.Escape("license"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("edition"), RegexOptions.IgnoreCase))
                            {
                                license.Add(Tuple.Create(node.Attributes[0].Value, text));
                            }
                            else if (Regex.IsMatch(text, Regex.Escape("freez"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("crash"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("timeout"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("block"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("run"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("later Analysis"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("later"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("analy"), RegexOptions.IgnoreCase) || Regex.IsMatch(text, Regex.Escape("time"), RegexOptions.IgnoreCase))
                            {
                                freeze.Add(Tuple.Create(node.Attributes[0].Value, text));
                            }
                            else
                            {
                                others.Add(Tuple.Create(node.Attributes[0].Value, text));
                            }
                        }

                    }

                    var assembly = Assembly.GetExecutingAssembly();
                    var reader = assembly.GetManifestResourceStream("Execute_Selected_Tests.Files.Ignored List.xlsx");
                    var file = Path.Combine(Environment.CurrentDirectory, "Ignored Tests Report.xlsx");
                    KillAllProcessesByName("EXCEL");
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }
                    using (Stream s = File.Create(file))
                    {
                        reader.CopyTo(s);
                    }
                    var xlApplication = new Microsoft.Office.Interop.Excel.Application
                    {
                        Visible = false
                    };

                    var workbook = xlApplication.Workbooks.Open(file);
                    var sheet = workbook.Worksheets[1];
                    for (var count = 0; count < sentForReview.Count; count++)
                    {
                        sheet.Cells[count + 2, 1] = count + 1;
                        sheet.Cells[count + 2, 2] = sentForReview[count].Item1;
                        sheet.Cells[count + 2, 3] = sentForReview[count].Item2;
                    }
                    sheet = workbook.Worksheets[2];
                    for (var count = 0; count < update.Count; count++)
                    {
                        sheet.Cells[count + 2, 1] = count + 1;
                        sheet.Cells[count + 2, 2] = update[count].Item1;
                        sheet.Cells[count + 2, 3] = update[count].Item2;
                    }
                    sheet = workbook.Worksheets[3];
                    for (var count = 0; count < license.Count; count++)
                    {
                        sheet.Cells[count + 2, 1] = count + 1;
                        sheet.Cells[count + 2, 2] = license[count].Item1;
                        sheet.Cells[count + 2, 3] = license[count].Item2;
                    }
                    sheet = workbook.Worksheets[4];
                    for (var count = 0; count < freeze.Count; count++)
                    {
                        sheet.Cells[count + 2, 1] = count + 1;
                        sheet.Cells[count + 2, 2] = freeze[count].Item1;
                        sheet.Cells[count + 2, 3] = freeze[count].Item2;
                    }
                    sheet = workbook.Worksheets[5];
                    for (var count = 0; count < test.Count; count++)
                    {
                        sheet.Cells[count + 2, 1] = count + 1;
                        sheet.Cells[count + 2, 2] = test[count].Item1;
                        sheet.Cells[count + 2, 3] = test[count].Item2;
                    }
                    sheet = workbook.Worksheets[6];
                    for (var count = 0; count < incomplete.Count; count++)
                    {
                        sheet.Cells[count + 2, 1] = count + 1;
                        sheet.Cells[count + 2, 2] = incomplete[count].Item1;
                        sheet.Cells[count + 2, 3] = incomplete[count].Item2;
                    }
                    sheet = workbook.Worksheets[7];
                    for (var count = 0; count < others.Count; count++)
                    {
                        sheet.Cells[count + 2, 1] = count + 1;
                        sheet.Cells[count + 2, 2] = others[count].Item1;
                        sheet.Cells[count + 2, 3] = others[count].Item2;
                    }
                    workbook.Save();
                    xlApplication.Visible = true;


                    Invoke(new Action(() => { button7.Enabled = true; }));
                    Invoke(new Action(() => { button7.Text = defaultText; }));
                    Invoke(new Action(() => { button6.Enabled = true; }));
                    Invoke(new Action(() => { button8.Enabled = true; }));
                }

                else
                {
                    MessageBox.Show(
                        "This Feature works with XML File only currently. \nPlease Select XML Results file to generate report");
                }
            }
        }

        private BackgroundWorker timeReportWorker;
        private void button8_Click(object sender, EventArgs e)
        {
            
            timeReportWorker = new BackgroundWorker();
            timeReportWorker.DoWork += (obj, ea) =>GetTimeReport();
            timeReportWorker.RunWorkerAsync();
        }


        private void GetTimeReport()
        {
            if (fileToOpen.Equals(string.Empty))
            {
                MessageBox.Show("Browse a file before clicking Generate.");
            }
            else
            {
                Invoke(new Action(() => { button8.Enabled = false; }));
                Invoke(new Action(() => { button6.Enabled = false; }));
                Invoke(new Action(() => { button7.Enabled = false; }));
                var defaultText = string.Empty;
                Invoke(new Action(() => { defaultText = button8.Text; }));
                Invoke(new Action(() => { button8.Text = "Please Wait.."; }));
                var timeList = new List<Tuple<string, double>>();
                var info = new FileInfo(fileToOpen);
                if (info.Extension.Equals(".xml"))
                {
                    var xmlDoc = new XmlDocument();
                    xmlDoc.Load(fileToOpen);

                    var testCaseNodes = xmlDoc.GetElementsByTagName("test-case");
                    foreach (XmlNode node in testCaseNodes)
                    {
                        if (!node.Attributes[3].Value.Equals("Ignored"))
                        {
                            timeList.Add(Tuple.Create(node.Attributes[0].Value, Convert.ToDouble(node.Attributes[5].Value) / 60));
                        }
                    }
                    timeList = timeList.OrderByDescending(x => x.Item2).ToList();
                    var assembly = Assembly.GetExecutingAssembly();
                    var reader = assembly.GetManifestResourceStream("Execute_Selected_Tests.Files.Time Report.xlsx");
                    var file = Path.Combine(Environment.CurrentDirectory, "Time Report.xlsx");
                    KillAllProcessesByName("EXCEL");
                    if (File.Exists(file))
                    {
                        File.Delete(file);
                    }
                    using (Stream s = File.Create(file))
                    {
                        reader.CopyTo(s);
                    }
                    var xlApplication = new Microsoft.Office.Interop.Excel.Application
                    {
                        Visible = false
                    };
                    var workbook = xlApplication.Workbooks.Open(file);
                    var sheet = workbook.Worksheets[1];
                    for (var count = 0; count < timeList.Count; count++)
                    {
                        sheet.Cells[count + 2, 1] = count + 1;
                        sheet.Cells[count + 2, 2] = timeList[count].Item1;
                        sheet.Cells[count + 2, 3] = timeList[count].Item2;
                    }

                    workbook.Save();
                    xlApplication.Visible = true;

                    Invoke(new Action(() => { button8.Enabled = true; }));
                    Invoke(new Action(() => { button8.Text = defaultText; }));
                    Invoke(new Action(() => { button6.Enabled = true; }));
                    Invoke(new Action(() => { button7.Enabled = true; }));

                }
                else
                {
                    MessageBox.Show(
                        "This Feature works with XML File only currently. \nPlease Select XML Results file to generate report");
                }
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            var file = Path.Combine(textBox3.Text, comboBox3.SelectedItem.ToString());

            if (File.Exists(file))
            {
                label111.ForeColor = Color.Green;
                label111.Text = "OK";
                button5.Enabled = true;
                button2.Enabled = true;
            }
            else
            {
                label111.ForeColor = Color.Red;
                label111.Text = $"Path does not contain\n {comboBox3.SelectedItem.ToString()}.";
                button5.Enabled = false;
                button2.Enabled = false;
            }
        }
    }
}
