using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using System.Data.OleDb;
using System.Threading;
namespace 数据转换
{
    class Program
    {
        public static readonly Fund fund = new Fund();
        static void Main(string[] args)
        {
            Console.WriteLine("基金字段检查程序说明：");
            Console.WriteLine("  存放匹配基金的字典文件为dict/dict.txt;");
            Console.WriteLine("  下载格式为Excel的输入文件,请存放到CSV文件下,只识别data.csv");
            Console.WriteLine("  下载格式为Txt的输入文件,请存放在TXT文件下,可以识别所有.txt文件");
            Console.WriteLine("  输出文件为OUTPUT");
            Console.WriteLine();
            Console.WriteLine("基金字段检查开始：");
            try
            {
                fund.FundNames = InitialFundNames("dict/dict.txt");
            }
            catch (Exception)
            {
                Console.WriteLine("系统找不到存放匹配基金的字典文件");
                Thread.Sleep(5000);
                return;
            }
            Console.WriteLine("请输入基金编号的开始字符：如 7");

            string startwith = Console.ReadLine();
            #region CSV数据处理
            int resultcount = 0, noresultcount = 0, index = 1;
            if (File.Exists("csv/data.csv"))
            {
                StreamReader reader = new StreamReader("csv/data.csv");
                StreamWriter resultWriter = new StreamWriter("output/result_csv.csv", false);
                StreamWriter noresultWriter = new StreamWriter("output/noresult_csv.csv", false);
                string line = reader.ReadLine();

                while ((line = reader.ReadLine()) != null)
                {
                    string[] accessionNumberandFund = fund.GetAccessionNumberAndFund(line);
                    if (accessionNumberandFund != null && fund.IsContain(fund.FundNames, startwith, accessionNumberandFund[1].Replace("\"", "")))
                    {
                        //匹配
                        resultcount++;
                        resultWriter.WriteLine(line);
                        Console.WriteLine("{0} {1} [满足]", index++, accessionNumberandFund[0]);

                    }
                    else
                    {
                        noresultcount++;
                        index++;
                        noresultWriter.WriteLine(line);
                    }
                }
                resultWriter.Flush();
                noresultWriter.Flush();

                resultWriter.Close();
                noresultWriter.Close();
                reader.Close();

                Thread.Sleep(1000);
                Console.WriteLine();

            }
            else
            {
                Console.WriteLine("CSV无数据");
            }
            #endregion

            #region TXT数据处理
            DirectoryInfo directorInfo = new DirectoryInfo("TXT");
            FileInfo[] fileInfos = directorInfo.GetFiles("*.txt");
            StreamWriter resultWriter1 = new StreamWriter("OUTPUT/result_txt.csv", false);
            StreamWriter noresultWriter1 = new StreamWriter("OUTPUT/noresult_txt.csv", false);
            int count_txt = 0;
            int count_no_txt = 0;
            int sum = 0;
            if (fileInfos != null && fileInfos.Length > 1)
            {
                foreach (var fileinfo in fileInfos)
                {
                    List<string> fundings = GetFunding(fileinfo.FullName);
                    sum += fundings.Count;
                    foreach (var fd in fundings)
                    {
                        string[] accessionNumberandFund = fund.GetAccessionNumberAndFund(fd);
                        if (accessionNumberandFund != null && fund.IsContain(fund.FundNames, startwith, accessionNumberandFund[1].Replace("\"", "")))
                        {
                            Console.WriteLine("{0}[满足]", accessionNumberandFund[0]);
                            resultWriter1.WriteLine(fd);
                            count_txt++;
                        }
                        else
                        {
                            noresultWriter1.WriteLine(fd);
                            count_no_txt++;
                        }

                    }
                }

                Thread.Sleep(1000);

            }
            else
            {
                Console.WriteLine("TXT没有文件");
            }
            resultWriter1.Flush();
            resultWriter1.Close();

            noresultWriter1.Flush();
            noresultWriter1.Close();
            #endregion
            Console.WriteLine();
            Console.WriteLine("CSV总共记录：{0}\r\n满足要求：{1}     \r\n不满足要求：{2}", index - 1, resultcount, noresultcount);
            Console.WriteLine();
            Console.WriteLine("TXT总共记录：{0}\r\n满足要求：{1}     \r\n不满足要求：{2}", sum, count_txt, count_no_txt);
            Console.WriteLine();
            Thread.Sleep(2000);
            Console.WriteLine("给个赞好不？");


            Console.ReadKey();
        }
        static Dictionary<string, bool> InitialFundNames(string filename)
        {
            StreamReader reader = new StreamReader(filename);
            string line = string.Empty;
            Dictionary<string, bool> fundnames = new Dictionary<string, bool>();
            while ((line = reader.ReadLine()) != null)
            {
                fundnames[line.Trim().ToLower()] = true;
            }
            reader.Close();
            reader.Dispose();
            return fundnames;
        }
        static void Foundation()
        {
            string reg = "NATIONAL NATURAL SCIENCE FOUNDATION OF CHINA\\s*\\[[\\s\\S]*?]|NATURAL SCIENCE FOUNDATION OF CHINA\\s*\\[[\\s\\S]*?]|NATIONAL SCIENCE FOUNDATION OF CHINA\\s*\\[[\\s\\S]*?]|NSFC\\s*\\[[\\s\\S]*?]|NATIONAL NATURAL SCIENCE FOUNDATION OF CHINA NSFC\\s*\\[[\\s\\S]*?]|NATIONAL NATURE SCIENCE FOUNDATION OF CHINA\\s*\\[[\\s\\S]*?]|NATIONAL SCIENCE FOUNDATION\\s*\\[[\\s\\S]*?]|NATIONAL SCIENCE FOUNDATION OF CHINA NSFC\\s*\\[[\\s\\S]*?]|NATIONAL NATURAL SCIENCES FOUNDATION OF CHINA\\s*\\[[\\s\\S]*?]|CHINESE NATIONAL NATURAL SCIENCE FOUNDATION\\s*\\[[\\s\\S]*?]|NATIONAL NATURAL SCIENCE FOUNDATIONS OF CHINA\\s*\\[[\\s\\S]*?]|KEY NATIONAL NATURAL SCIENCE FOUNDATION OF CHINA\\s*\\[[\\s\\S]*?]|NATURAL SCIENCES FOUNDATION OF CHINA\\s*\\[[\\s\\S]*?]|CHINA NATURAL SCIENCE FOUNDATION\\s*\\[[\\s\\S]*?]|NATURAL SCIENCE FOUNDATION OF CHINA NSFC\\s*\\[[\\s\\S]*?]|CHINA NATIONAL SCIENCE FOUNDATION\\s*\\[[\\s\\S]*?]|NATIONAL SCIENCE FOUNDATION CHINA\\s*\\[[\\s\\S]*?]|NATURE SCIENCE FOUNDATION OF CHINA\\s*\\[[\\s\\S]*?]|NATIONAL SCIENTIFIC FOUNDATION OF CHINA\\s*\\[[\\s\\S]*?]|CHINESE NATIONAL SCIENCE FOUNDATION\\s*\\[[\\s\\S]*?]|National Natural Science Foundation of China (NSFC)\\s*\\[[\\s\\S]*?]|NSFC (National Natural Science Foundation of China)\\s*\\[[\\s\\S]*?]|China National Natural Science Foundation\\s*\\[[\\s\\S]*?]|Natural Science Foundation of China (NSFC)\\s*\\[[\\s\\S]*?]|National Science Foundation of P. R. China\\s*\\[[\\s\\S]*?]";
            // string reg = "NATIONAL NATURAL SCIENCE FOUNDATION OF CHINA |NATURAL SCIENCE FOUNDATION OF CHINA |NATIONAL SCIENCEFOUNDATION OF CHINA |NSFC|NATIONAL NATURAL SCIENCE FOUNDATION OF CHINA NSFC |NATIONAL NATURE SCIENCE FOUNDATION OF CHINA |NATIONAL SCIENCE FOUNDATION |NATIONAL SCIENCE FOUNDATION OF CHINA NSFC |NATIONAL NATURAL SCIENCES FOUNDATION OF CHINA |CHINESE NATIONAL NATURAL SCIENCE FOUNDATION |NATIONAL NATURAL SCIENCE FOUNDATIONS OF CHINA |KEY NATIONAL NATURAL SCIENCE FOUNDATION OF CHINA |NATURAL SCIENCES FOUNDATION OF CHINA |CHINA NATURAL SCIENCE FOUNDATION |NATURAL SCIENCE FOUNDATION OF CHINA NSFC |CHINA NATIONAL SCIENCE FOUNDATION |NATIONAL SCIENCE FOUNDATION CHINA |NATURE SCIENCE FOUNDATION OF CHINA |NATIONAL SCIENTIFIC FOUNDATION OF CHINA |CHINESE NATIONAL SCIENCE FOUNDATION |National Natural Science Foundation of China (NSFC)|NSFC (National Natural Science Foundation of China)|China National Natural Science Foundation|Natural Science Foundation of China (NSFC) ";
            Regex regex = new Regex(reg, RegexOptions.IgnoreCase);
            StreamReader reader = new StreamReader("data.csv", Encoding.Default);

            string line = string.Empty;
            StringBuilder result = new StringBuilder();
            StringBuilder noresult = new StringBuilder();
            int ii = 0;
            while ((line = reader.ReadLine()) != null)
            {
                ii++;
                if (string.IsNullOrEmpty(line))
                {
                    line = reader.ReadLine();
                    continue;
                }
                if (regex.IsMatch(line))
                {
                    bool isTrue = false;
                    MatchCollection collection = regex.Matches(line);
                    int start = 0;
                    foreach (Match match in collection)
                    {
                        string value = match.Value;

                        start = line.IndexOf(value, start);
                        start = start - 1;

                        //start = line.IndexOf('[', start);
                        //int end = line.IndexOf(']', start);
                        //string lines = line.Substring(start, end - start + 1);
                        //string[] words = lines.Split(new string[] { ",", "[", "]", "NSFC" }, StringSplitOptions.RemoveEmptyEntries);
                        string[] words = value.Split(new string[] { ",", "[", "]", "NSFC" }, StringSplitOptions.RemoveEmptyEntries);

                        for (int i = 0; i < words.Length; i++)
                        {
                            if (words[i].Trim().StartsWith("7"))
                            {
                                isTrue = true;
                                Console.WriteLine(words[i]);
                                continue;
                            }
                        }


                    }
                    if (isTrue)
                    {
                        result.AppendLine(line);
                    }
                    else
                    {
                        noresult.AppendLine(line);
                    }

                }
                else
                {
                    noresult.AppendLine(line);
                }
            }
            StreamWriter writer = new StreamWriter("result.csv", false, Encoding.Default);
            writer.Write(result);
            writer.Flush();
            writer.Close();
            reader.Close();

            StreamWriter writer1 = new StreamWriter("noresult.csv", false, Encoding.Default);
            writer1.Write(noresult);
            writer1.Flush();
            writer1.Close();

        }
        static void Category()
        {

            OleDbConnection connection = new OleDbConnection();
            connection.ConnectionString = string.Format("Provider=Microsoft.Ace.OleDb.12.0;Data Source='{0}';Extended Properties='Excel 12.0;HDR=YES'", "category.xlsx"); ;
            OleDbCommand cmd = new OleDbCommand();
            cmd.CommandText = "select Issn,Categorys from [Category$] order by [Issn]";
            cmd.Connection = connection;
            connection.Open();
            IDataReader iReader = cmd.ExecuteReader();
            string tempIssn = string.Empty;
            StringBuilder result = new StringBuilder();
            string issn = string.Empty;
            while (iReader.Read())
            {
                issn = iReader["Issn"].ToString();
                if (string.IsNullOrEmpty(issn))
                    continue;
                if (issn == tempIssn)
                {
                    result.Append(";");
                    result.Append(iReader["Categorys"].ToString());
                }
                else
                {
                    tempIssn = iReader["Issn"].ToString();
                    result.AppendLine();
                    result.Append(tempIssn);
                    result.Append("~");
                    result.Append(iReader["Categorys"].ToString());

                }

            }
            StreamWriter writer1 = new StreamWriter("category.csv", false, Encoding.Default);
            writer1.Write(result);
            writer1.Flush();
            writer1.Close();
            connection.Close();
        }
        static string wosnumberReg = "WOS:[0-9]*";
        static bool IsContains(Dictionary<string, bool> fundsname, string startswith, string line)
        {
            Match match = Regex.Match(line, wosnumberReg);
            if (match == null)
            {
                return false;
            }
            string wownumber = match.Value;

            string content = line.Substring(wownumber.Length + 1);
            //使用分号隔开
            string[] funds = content.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < funds.Length; i++)
            {
                string[] temp = funds[i].Split(new char[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
                if (temp.Length == 0)
                    continue;
                bool isMatchName = false;
                if (temp[0].Contains('('))
                {
                    string[] temp1 = temp[0].Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < temp1.Length; j++)
                    {
                        if (fundsname.ContainsKey(temp1[j].Trim()))
                        {
                            isMatchName = true;
                            break;
                        }
                    }
                }
                else if (fundsname.ContainsKey(temp[0].Trim()))
                {
                    isMatchName = true;
                }
                if (isMatchName)
                {
                    string[] temp3 = temp[1].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int h = 0; h < temp3.Length; h++)
                    {
                        if (temp3[i].TrimStart().StartsWith(startswith))
                        {
                            return true;
                        }
                    }
                }
                else
                    continue;

            }
            return false;
        }
        static List<string> GetFunding(string file)
        {
            string line = string.Empty;
            string ut = string.Empty;
            string mark = string.Empty;
            StringBuilder funding = new StringBuilder();
            List<string> fundingList = new List<string>();
            using (StreamReader reader = new StreamReader(file))
            {
                while ((line = reader.ReadLine()) != null)
                {

                    if (line.StartsWith("FU"))
                    {
                        funding.Append(line.Substring(3));
                        mark = "FU";
                        continue;
                    }
                    else if (line.StartsWith(" ") && mark == "FU")
                    {
                        funding.Append(line.Substring(2));
                        continue;
                    }
                    else
                    {
                        mark = string.Empty;
                    }

                    if (line.StartsWith("UT"))
                    {
                        ut = line.Substring(3);
                        continue;

                    }
                    if (line.StartsWith("ER"))
                    {
                        fundingList.Add(string.Format("{0},\"{1}\"", ut, funding.ToString()));
                        ut = string.Empty;
                        funding.Clear();
                    }
                }
            }
            return fundingList;
        }
    }
    class Fund
    {
        public Dictionary<string, bool> FundNames { set; get; }
        public string StartWith { set; get; }
        public readonly string AccessionNumberReg = "WOS:[0-9]*";
        public bool IsContain(Dictionary<string, bool> fundnames, string startwith, string fundline)
        {
            //使用分号隔开
            string[] funds = fundline.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < funds.Length; i++)
            {
                string[] temp = funds[i].Split(new char[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
                if (temp.Length == 0)
                    continue;
                bool isMatchName = false;
                if (temp[0].Contains('('))
                {
                    string[] temp1 = temp[0].Split(new char[] { '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
                    for (int j = 0; j < temp1.Length; j++)
                    {
                        if (fundnames.ContainsKey(temp1[j].Trim().ToLower()))
                        {
                            isMatchName = true;
                            break;
                        }
                    }
                }
                else if (fundnames.ContainsKey(temp[0].Trim().ToLower()))
                {
                    isMatchName = true;
                }
                if (isMatchName)
                {
                    if (temp.Length > 1)
                    {
                        string[] temp3 = temp[1].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        for (int h = 0; h < temp3.Length; h++)
                        {
                            if (temp3[h].ToLower().TrimStart().StartsWith(startwith))
                            {
                                return true;
                            }
                        }
                    } continue;
                }
                else
                    continue;

            }
            return false;
        }
        public string[] GetAccessionNumberAndFund(string line)
        {
            Match match = Regex.Match(line, AccessionNumberReg);
            if (match == null)
            {
                return null;
            }
            string wownumber = match.Value;
            string content = line.Substring(wownumber.Length + 1);
            return new string[] { wownumber, content };
        }
    }

}
