using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace SmartSchool.Evaluation.Reports
{
    public partial class SemesterScoreReportFormNew : SelectSemesterForm
    {
        private Dictionary<string, List<string>> _userType = new Dictionary<string, List<string>>();
        private MemoryStream _defaultTemplate = new MemoryStream(Properties.Resources.�Ǵ����Z��New);
        private MemoryStream _template = null;
        private bool _useDefaultTemplate = false;
        private int _receiver = 0;
        private int _address = 0;
        private bool _over100 = false;
        private byte[] _buffer = null;
        private string _resitSign = "*";
        private string _repeatSign = "#";

        public Dictionary<string, List<string>> UserDefinedType
        {
            get { return _userType; }
        }

        public MemoryStream Template
        {
            get
            {
                if (_useDefaultTemplate)
                    return _defaultTemplate;
                else if (_template != null)
                    return _template;
                else
                    return _defaultTemplate;
            }
        }
        public int Receiver
        {
            get { return _receiver; }
        }
        public int Address
        {
            get { return _address; }
        }

        public string ResitSign
        {
            get { return _resitSign; }
        }
        public string RepeatSign
        {
            get { return _repeatSign; }
        }

        public bool AllowMoralScoreOver100
        {
            get { return _over100; }
        }


        public SemesterScoreReportFormNew()
        {
            InitializeComponent();
            LoadPreference();
        }

        private void LoadPreference()
        {
            XmlElement config = CurrentUser.Instance.Preference["�Ǵ����Z��New"];
            if (config != null)
            {
                //�ϥΪ̳]�w�����O
                _userType.Clear();

                foreach (XmlElement type in config.SelectNodes("Type"))
                {
                    string typeName = type.GetAttribute("Text");

                    if (!_userType.ContainsKey(typeName))
                        _userType.Add(typeName, new List<string>());

                    foreach (XmlElement absence in type.SelectNodes("Absence"))
                    {
                        string absenceName = absence.GetAttribute("Text");

                        if (!_userType[typeName].Contains(absenceName))
                            _userType[typeName].Add(absenceName);
                    }
                }

                //�d��
                if (config.HasAttribute("UseDefault"))
                    _useDefaultTemplate = bool.Parse(config.GetAttribute("UseDefault"));
                else
                {
                    config.SetAttribute("UseDefault", "True");
                    CurrentUser.Instance.Preference["�Ǵ����Z��New"] = config;
                }

                XmlElement customize = (XmlElement)config.SelectSingleNode("CustomizeTemplate");

                if (customize != null)
                {
                    if (!string.IsNullOrEmpty(customize.InnerText))
                    {
                        string templateBase64 = customize.InnerText;
                        _buffer = Convert.FromBase64String(templateBase64);
                        _template = new MemoryStream(_buffer);
                    }
                }
                else
                {
                    XmlElement newCustomize = config.OwnerDocument.CreateElement("CustomizeTemplate");
                    config.AppendChild(newCustomize);
                    CurrentUser.Instance.Preference["�Ǵ����Z��New"] = config;
                }

                //�C�L��T
                XmlElement print = (XmlElement)config.SelectSingleNode("Print");
                if (print != null)
                {
                    _receiver = int.Parse(print.GetAttribute("Name"));
                    _address = int.Parse(print.GetAttribute("Address"));
                    _resitSign = print.GetAttribute("ResitSign");
                    _repeatSign = print.GetAttribute("RepeatSign");
                    if (!string.IsNullOrEmpty(print.GetAttribute("AllowMoralScoreOver100")))
                        _over100 = bool.Parse(print.GetAttribute("AllowMoralScoreOver100"));
                }
                else
                {
                    XmlElement newPrint = config.OwnerDocument.CreateElement("Print");
                    newPrint.SetAttribute("Name", "0");
                    newPrint.SetAttribute("Address", "0");
                    newPrint.SetAttribute("ResitSign", "*");
                    newPrint.SetAttribute("RepeatSign", "#");
                    newPrint.SetAttribute("AllowMoralScoreOver100", "False");
                    _receiver = 0;
                    _address = 0;
                    _resitSign = "*";
                    _repeatSign = "#";
                    _over100 = false;
                    config.AppendChild(newPrint);
                    CurrentUser.Instance.Preference["�Ǵ����Z��New"] = config;
                }
            }
            else
            {
                #region ���ͪťճ]�w��
                config = new XmlDocument().CreateElement("�Ǵ����Z��New");
                CurrentUser.Instance.Preference["�Ǵ����Z��New"] = config;
                #endregion
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SemesterScoreReportConfigNew configForm = new SemesterScoreReportConfigNew(_useDefaultTemplate, _buffer, _receiver, _address, _resitSign, _repeatSign, _over100);
            if (configForm.ShowDialog() == DialogResult.OK)
            {
                LoadPreference();
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SelectTypeForm form = new SelectTypeForm("�Ǵ����Z��New");
            if (form.ShowDialog() == DialogResult.OK)
            {
                LoadPreference();
            }
        }
    }
}