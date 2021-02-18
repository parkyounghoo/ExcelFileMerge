using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ExcelFileMerge
{
    public partial class Form1 : Form
    {
        private Timer _timer;
        private OpenFileDialog openFile1;
        private OpenFileDialog openFile2;
        private int point;
        private int timerCnt;
        private string fileName;
        private string safeFileName;
        private long fileSize1;
        private long fileSize2;
        private int filesizeCnt;

        public Form1()
        {
            InitializeComponent();

            richTextBox1.ReadOnly = true;
            textBox1.Text = "";
            textBox1.ReadOnly = true;
            textBox2.Text = "";
            textBox2.ReadOnly = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFile1 = new OpenFileDialog();
            openFile1.Title = "파일을 선택하세요";
            openFile1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.csv";

            openFile1.ShowDialog();
            if (openFile1.FileNames.Length > 0)
            {
                safeFileName = openFile1.SafeFileName;
                foreach (string filename in openFile1.FileNames)
                {
                    this.textBox1.Text = filename;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFile2 = new OpenFileDialog();
            openFile2.Title = "파일을 선택하세요";
            openFile2.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm;*.csv";

            openFile2.ShowDialog();
            if (openFile2.FileNames.Length > 0)
            {
                foreach (string filename in openFile2.FileNames)
                {
                    this.textBox2.Text = filename;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (!textBox1.Text.Contains(""))
                {
                    MessageBox.Show("파일을 업로드해주세요");
                }
                else if (textBox2.Text == "")
                {
                    MessageBox.Show("파일을 업로드해주세요");
                }
                else
                {
                    FileInfo fInfo = new FileInfo(textBox1.Text);
                    fileSize1 = fInfo.Length;

                    FileStream fs = null;

                    fileName = textBox1.Text.Replace(safeFileName, "old_data_reservation_" + DateTime.Now.ToString("yyyyMMdd") + ".csv");

                    FileInfo fileDel = new FileInfo(fileName);
                    if (fileDel.Exists) // 삭제할 파일이 있는지
                    {
                        fileDel.Delete(); // 없어도 에러안남
                    }

                    point = 0;
                    timerCnt = 0;
                    filesizeCnt = 0;
                    _timer = new Timer();
                    _timer.Interval = 1000;
                    _timer.Tick += _timer_Tick;
                    _timer.Start();
                    richTextBox1.Text = "Loading.....";

                    string path = Environment.CurrentDirectory + @"\24Hour_Reservation_v2\24Hour_Reservation_v2.exe";
                    string path1 = @"""" + textBox1.Text + @"""";
                    string path2 = @"""" + textBox2.Text + @"""";
                    ProcessStartInfo startInfo = new ProcessStartInfo(path, path1 + " " + path2);

                    Process process = new Process();

                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;

                    startInfo.CreateNoWindow = true;

                    process.StartInfo = startInfo;

                    process.Start();

                    process.Close();
                }
            }
            catch (Exception ex)
            {
                richTextBox1.Text = fileName + " 파일이 사용중 입니다. 프로그램 종료 후 다시 실행하세요.";
            }
        }

        private void _timer_Tick(object sender, EventArgs e)
        {
            button3.Enabled = true;
            point++;
            timerCnt++;
            string txt = "";
            for (int i = 0; i < point; i++)
            {
                txt += ".";
            }

            if (point == 6)
            {
                point = 0;
            }

            richTextBox1.Text = "Loading" + txt;

            FileInfo fileInfo = new FileInfo(fileName);
            if (fileInfo.Exists)
            {
                fileSize2 = fileInfo.Length;

                if (fileSize1 < fileSize2)
                {
                    filesizeCnt++;
                    if (filesizeCnt == 3)
                    {
                        //다시 파일을 닫아줘야 한다.
                        richTextBox1.Text = "작업이 완료되었습니다. (결과 파일 경로 : " + fileName + ")";
                        _timer.Stop();
                    }   
                }
            }
        }

        private bool fileCheck()
        {
            FileStream fs = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open,
                        FileAccess.Read, FileShare.None);

                if (fs != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }

                fs.Close();
            }
            catch (IOException)
            {
                fs.Close();
                return false;
            }
        }
    }
}