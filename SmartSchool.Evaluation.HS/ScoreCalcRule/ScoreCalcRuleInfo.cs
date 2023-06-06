using FISCA.DSAUtil;
using SmartSchool.StudentRelated;
using SmartSchool.TagManage;
using System.Collections.Generic;
using System.Xml;

namespace SmartSchool.Evaluation.ScoreCalcRule
{
    public class ScoreCalcRuleInfo
    {
        private readonly string _ID;
        private readonly string _Name;
        private readonly XmlElement _ScoreCalcRuleElement;

        private readonly string _SchoolYear;
        private readonly string _TrimName;
        //private readonly bool _DefinedSubjectInfoByGPlan;

        /// <summary>
        /// ���Z�p��W�h
        /// </summary>
        /// <param name="scrElement"></param>
        internal ScoreCalcRuleInfo(XmlElement scrElement)
        {
            _ID = scrElement.GetAttribute("ID");
            _Name = scrElement.SelectSingleNode("Name").InnerText;
            _ScoreCalcRuleElement = (XmlElement)scrElement.SelectSingleNode("Content/ScoreCalcRule");

            _SchoolYear = _ScoreCalcRuleElement.HasAttribute("SchoolYear") ? _ScoreCalcRuleElement.GetAttribute("SchoolYear") : string.Empty;
            _TrimName = _Name;
            if (!string.IsNullOrEmpty(_SchoolYear))
            {
                //_TrimName = _Name.TrimStart(_SchoolYear.ToCharArray());
                _TrimName = _Name.Substring(_SchoolYear.Length);
            }
            //_DefinedSubjectInfoByGPlan = false;
            //if (_ScoreCalcRuleElement.SelectSingleNode("�Ǥ��έ׽Ҹ�T�ĭp�覡") == null || ((XmlElement)_ScoreCalcRuleElement.SelectSingleNode("�Ǥ��έ׽Ҹ�T�ĭp�覡")).GetAttribute("�ѽҵ{�W������o") == "True")
            //    _DefinedSubjectInfoByGPlan = true;
        }

        public string ID { get { return _ID; } }
        public string Name { get { return _Name; } }
        public string SchoolYear { get { return _SchoolYear; } }
        public string TrimName { get { return _TrimName; } }
        //public bool DefinedSubjectInfoByGPlan{get { return _DefinedSubjectInfoByGPlan; }}
        public XmlElement ScoreCalcRuleElement
        {
            get
            {
                return (XmlElement)(new XmlDocument().ImportNode(_ScoreCalcRuleElement, true));
            }
        }

        public XmlElement CalculateSemesterEntryScore(XmlElement semesterSubjectScore)
        {
            Dictionary<string, decimal> entryCreditCount = new Dictionary<string, decimal>();
            Dictionary<string, List<decimal>> entrySubjectScores = new Dictionary<string, List<decimal>>();
            Dictionary<string, decimal> entryDividend = new Dictionary<string, decimal>();
            Dictionary<string, bool> calcEntry = new Dictionary<string, bool>();
            Dictionary<string, bool> calcInStudy = new Dictionary<string, bool>();
            List<string> takeScore = new List<string>();
            bool takeRepairScore = false;
            //��Ǧ��
            int decimals = 2;
            //�i��Ҧ�
            WearyDogComputer.RoundMode mode = WearyDogComputer.RoundMode.�|�ˤ��J;
            //���Z�~�Ťέp��W�h�Ҧs�b�A���\�p�⦨�Z
            bool canCalc = true;
            #region ���o���Z�~�Ÿ�p��W�h

            #region �B�z�p��W�h
            XmlElement scoreCalcRule = (XmlElement)_ScoreCalcRuleElement;//ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID).ScoreCalcRuleElement;

            DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
            bool tryParsebool;
            int tryParseint;
            decimal tryParseDecimal;

            #region ��Ǧ��
            if (scoreCalcRule.SelectSingleNode("�U�����Z�p����/�Ǵ��������Z�p����") != null)
            {
                if (int.TryParse(helper.GetText("�U�����Z�p����/�Ǵ��������Z�p����/@���"), out tryParseint))
                    decimals = tryParseint;
                if (bool.TryParse(helper.GetText("�U�����Z�p����/�Ǵ��������Z�p����/@�|�ˤ��J"), out tryParsebool) && tryParsebool)
                    mode = WearyDogComputer.RoundMode.�|�ˤ��J;
                if (bool.TryParse(helper.GetText("�U�����Z�p����/�Ǵ��������Z�p����/@�L����˥h"), out tryParsebool) && tryParsebool)
                    mode = WearyDogComputer.RoundMode.�L����˥h;
                if (bool.TryParse(helper.GetText("�U�����Z�p����/�Ǵ��������Z�p����/@�L����i��"), out tryParsebool) && tryParsebool)
                    mode = WearyDogComputer.RoundMode.�L����i��;
            }
            #endregion
            #region �p�����O
            foreach (string entry in new string[] { "��|", "�Ƿ~", "�꨾�q��", "���d�P�@�z", "��߬��", "�M�~���" })
            {
                if (scoreCalcRule.SelectSingleNode("�������Z�p�ⶵ��") == null || scoreCalcRule.SelectSingleNode("�������Z�p�ⶵ��/" + entry) == null || ((XmlElement)scoreCalcRule.SelectSingleNode("�������Z�p�ⶵ��/" + entry)).GetAttribute("�p�⦨�Z") == "True")
                    calcEntry.Add(entry, true);
                else
                    calcEntry.Add(entry, false);

                // 2014/3/20�A�ק��S���Ŀ� �w�]�֤J�Ǵ��Ƿ~���Z ChenCT
                if (scoreCalcRule.SelectSingleNode("�������Z�p�ⶵ��") == null || scoreCalcRule.SelectSingleNode("�������Z�p�ⶵ��/" + entry) == null || ((XmlElement)scoreCalcRule.SelectSingleNode("�������Z�p�ⶵ��/" + entry)).GetAttribute("�֤J�Ǵ��Ƿ~���Z") == "True")
                    calcInStudy.Add(entry, true);
                else
                    calcInStudy.Add(entry, false);
            }
            #endregion
            #region �ĭp���Z���
            // �B�z�ɭצ��Z
            bool.TryParse(helper.GetText("�������Z�p��ĭp���Z���/@�ɭצ��Z"), out takeRepairScore);

            foreach (string item in new string[] { "��l���Z", "�ɦҦ��Z", "���צ��Z", "���u�ĭp���Z", "�Ǧ~�վ㦨�Z" })
            {
                if (!bool.TryParse(helper.GetText("�������Z�p��ĭp���Z���/@" + item), out tryParsebool) || tryParsebool)
                {//�S���]�w�o�����Z�]�w�W�h(�w�]true)�Ϊ̳]�w�ȬOtrue
                    takeScore.Add(item);
                }
            }
            #endregion

            #endregion

            #endregion
            Dictionary<string, decimal> entryScores = new Dictionary<string, decimal>();

            #region �N���Z����U�������O��
            foreach (XmlNode subjectNode in semesterSubjectScore.SelectNodes("Subject"))
            {
                XmlElement subjectElement = (XmlElement)subjectNode;
                //���p�Ǥ��Τ��ݵ������κ�
                if (subjectElement.GetAttribute("���ݵ���") == "�O" || subjectElement.GetAttribute("���p�Ǥ�") == "�O")
                    continue;

                // �Y���ɭצ��Z�B�ɭצ��Z���ĭp�A�h���p��
                if (takeRepairScore == false)
                {
                    if (subjectElement.GetAttribute("�O�_�ɭצ��Z") == "�O")
                        continue;
                }

                string subjectCode = subjectElement.GetAttribute("�׽Ҭ�إN�X");
                if (subjectCode.Length >= 23) //�@23�X
                {
                    if (subjectCode[16].ToString() + subjectCode[18].ToString() == "9D" || subjectCode[16].ToString() + subjectCode[18].ToString() == "9d")
                        continue;
                }
                #region �������O��Ǥ���
                string entry = subjectElement.GetAttribute("�}�Ҥ������O");
                decimal credit = 0;
                decimal.TryParse(subjectElement.GetAttribute("�}�ҾǤ���"), out credit);
                #endregion
                decimal maxScore = 0;
                decimal original = 0;
                if (decimal.TryParse(subjectElement.GetAttribute("��l���Z"), out tryParseDecimal))
                    original = tryParseDecimal;
                #region ���o�̰�����
                foreach (var item in takeScore)
                {
                    if (decimal.TryParse(subjectElement.GetAttribute(item), out tryParseDecimal) && maxScore < tryParseDecimal)
                        maxScore = tryParseDecimal;
                }
                #endregion
                switch (entry)
                {
                    case "��|":
                    case "�꨾�q��":
                    case "���d�P�@�z":
                    case "��߬��":
                    case "�M�~���":
                        //�p��������Z
                        if (calcEntry[entry])
                        {
                            #region original
                            //�[�`�Ǥ���
                            if (!entryCreditCount.ContainsKey(entry + "(��l)"))
                                entryCreditCount.Add(entry + "(��l)", credit);
                            else
                                entryCreditCount[entry + "(��l)"] += credit;
                            //�[�J�N���Z��Ƥ���
                            if (!entrySubjectScores.ContainsKey(entry + "(��l)")) entrySubjectScores.Add(entry + "(��l)", new List<decimal>());
                            entrySubjectScores[entry + "(��l)"].Add(original);
                            //�[�v�`�p
                            if (!entryDividend.ContainsKey(entry + "(��l)"))
                                entryDividend.Add(entry + "(��l)", original * credit);
                            else
                                entryDividend[entry + "(��l)"] += (original * credit);
                            #endregion
                            #region maxScore
                            //�[�`�Ǥ���
                            if (!entryCreditCount.ContainsKey(entry))
                                entryCreditCount.Add(entry, credit);
                            else
                                entryCreditCount[entry] += credit;
                            //�[�J�N���Z��Ƥ���
                            if (!entrySubjectScores.ContainsKey(entry)) entrySubjectScores.Add(entry, new List<decimal>());
                            entrySubjectScores[entry].Add(maxScore);
                            //�[�v�`�p
                            if (!entryDividend.ContainsKey(entry))
                                entryDividend.Add(entry, maxScore * credit);
                            else
                                entryDividend[entry] += (maxScore * credit);
                            #endregion
                        }
                        //�N��ئ��Z�P�Ƿ~���Z�@�֭p��
                        if (calcInStudy[entry])
                        {
                            #region original
                            //�[�`�Ǥ���
                            if (!entryCreditCount.ContainsKey("�Ƿ~" + "(��l)"))
                                entryCreditCount.Add("�Ƿ~" + "(��l)", credit);
                            else
                                entryCreditCount["�Ƿ~" + "(��l)"] += credit;
                            //�[�J�N���Z��Ƥ���
                            if (!entrySubjectScores.ContainsKey("�Ƿ~" + "(��l)")) entrySubjectScores.Add("�Ƿ~" + "(��l)", new List<decimal>());
                            entrySubjectScores["�Ƿ~" + "(��l)"].Add(original);
                            //�[�v�`�p
                            if (!entryDividend.ContainsKey("�Ƿ~" + "(��l)"))
                                entryDividend.Add("�Ƿ~" + "(��l)", original * credit);
                            else
                                entryDividend["�Ƿ~" + "(��l)"] += (original * credit);
                            #endregion
                            #region maxScore
                            //�[�`�Ǥ���
                            if (!entryCreditCount.ContainsKey("�Ƿ~"))
                                entryCreditCount.Add("�Ƿ~", credit);
                            else
                                entryCreditCount["�Ƿ~"] += credit;
                            //�[�J�N���Z��Ƥ���
                            if (!entrySubjectScores.ContainsKey("�Ƿ~")) entrySubjectScores.Add("�Ƿ~", new List<decimal>());
                            entrySubjectScores["�Ƿ~"].Add(maxScore);
                            //�[�v�`�p
                            if (!entryDividend.ContainsKey("�Ƿ~"))
                                entryDividend.Add("�Ƿ~", maxScore * credit);
                            else
                                entryDividend["�Ƿ~"] += (maxScore * credit);
                            #endregion
                        }
                        break;

                    case "�Ƿ~":
                    default:
                        #region original
                        //�[�`�Ǥ���
                        if (!entryCreditCount.ContainsKey("�Ƿ~" + "(��l)"))
                            entryCreditCount.Add("�Ƿ~" + "(��l)", credit);
                        else
                            entryCreditCount["�Ƿ~" + "(��l)"] += credit;
                        //�[�J�N���Z��Ƥ���
                        if (!entrySubjectScores.ContainsKey("�Ƿ~" + "(��l)")) entrySubjectScores.Add("�Ƿ~" + "(��l)", new List<decimal>());
                        entrySubjectScores["�Ƿ~" + "(��l)"].Add(original);
                        //�[�v�`�p
                        if (!entryDividend.ContainsKey("�Ƿ~" + "(��l)"))
                            entryDividend.Add("�Ƿ~" + "(��l)", original * credit);
                        else
                            entryDividend["�Ƿ~" + "(��l)"] += (original * credit);
                        #endregion
                        #region maxScore
                        //�[�`�Ǥ���
                        if (!entryCreditCount.ContainsKey("�Ƿ~"))
                            entryCreditCount.Add("�Ƿ~", credit);
                        else
                            entryCreditCount["�Ƿ~"] += credit;
                        //�[�J�N���Z��Ƥ���
                        if (!entrySubjectScores.ContainsKey("�Ƿ~")) entrySubjectScores.Add("�Ƿ~", new List<decimal>());
                        entrySubjectScores["�Ƿ~"].Add(maxScore);
                        //�[�v�`�p
                        if (!entryDividend.ContainsKey("�Ƿ~"))
                            entryDividend.Add("�Ƿ~", maxScore * credit);
                        else
                            entryDividend["�Ƿ~"] += (maxScore * credit);
                        #endregion
                        break;
                }
            }
            #endregion

            XmlDocument doc = new XmlDocument();
            XmlElement entryScoreRoot = doc.CreateElement("SemesterEntryScore");
            #region �B�z�p��U�������O�����Z
            foreach (string entry in entryCreditCount.Keys)
            {
                decimal entryScore = 0;
                #region �p��entryScore
                if (entryCreditCount[entry] == 0)
                {
                    foreach (decimal score in entrySubjectScores[entry])
                    {
                        entryScore += score;
                    }
                    entryScore = (entryScore / entrySubjectScores[entry].Count);
                }
                else
                {
                    //�Υ[�v�`�����Ǥ���
                    entryScore = (entryDividend[entry] / entryCreditCount[entry]);
                }
                #endregion
                //��Ǧ�ƳB�z
                entryScore = WearyDogComputer.GetRoundScore(entryScore, decimals, mode);
                #region ��JXml
                XmlElement entryElement = doc.CreateElement("Entry");
                entryElement.SetAttribute("����", entry);
                entryElement.SetAttribute("���Z", entryScore.ToString());
                entryScoreRoot.AppendChild(entryElement);
                #endregion
            }
            #endregion
            return entryScoreRoot;
        }

        public decimal GetStudentPassScore(BriefStudentData student, int gradeYear)
        {
            string childElement;
            switch (gradeYear)
            {
                case 1: childElement = "�@�~�Ťή�з�"; break;
                case 2: childElement = "�G�~�Ťή�з�"; break;
                case 3: childElement = "�T�~�Ťή�з�"; break;
                case 4: childElement = "�|�~�Ťή�з�"; break;
                default: childElement = "�W�L�F��"; break;
            }
            decimal passScore = decimal.MaxValue, tryPraseScore;
            DSXmlHelper helper = new DSXmlHelper(_ScoreCalcRuleElement);
            foreach (XmlElement element in helper.GetElements("�ή�з�/�ǥ����O"))
            {
                string tagName = element.GetAttribute("���O");
                if (tagName == "�w�]" && decimal.TryParse(element.GetAttribute(childElement), out tryPraseScore))
                {
                    if (tryPraseScore < passScore)
                        passScore = tryPraseScore;
                }
                else
                {
                    foreach (TagInfo tag in student.Tags)
                    {
                        if (((tag.Prefix == "" && tagName == tag.Name) || tagName == tag.FullName) && decimal.TryParse(element.GetAttribute(childElement), out tryPraseScore))
                        {
                            if (tryPraseScore < passScore)
                                passScore = tryPraseScore;
                            break;
                        }
                    }
                }
            }
            if (passScore == decimal.MaxValue)
                passScore = 60;
            return passScore;
        }

        private enum RoundMode { �|�ˤ��J, �L����i��, �L����˥h }
        private decimal GetRoundScore(decimal entryScore, int decimals, SmartSchool.Evaluation.WearyDogComputer.RoundMode mode)
        {
            return WearyDogComputer.GetRoundScore(entryScore, decimals, mode);
        }
    }
}
