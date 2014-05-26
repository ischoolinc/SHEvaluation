namespace SchoolYearScoreReport
{
    //using SchoolYearScoreReport.Properties;
    using SmartSchool.Customization.Data;
    using System;
    using System.Collections.Generic;
    using System.Xml;

    public class Config
    {
        // Fields
        private bool _allow_over;
        private byte[] _custom_template;
        private bool _isUploadNewTemplate = false;
        private int _receive_address;
        private int _receive_name;
        private string _repeat_sign;
        private string _resit_sign;
        private decimal _school_year = 50M;
        private Dictionary<string, List<string>> _select_types = new Dictionary<string, List<string>>();
        private bool _use_default = true;

        // Methods
        public Config()
        {
            this.LoadSchoolInfo();
            this.LoadPreference();
        }

        private void LoadPreference()
        {
            XmlElement config = SmartSchool.Customization.Data.SystemInformation.Preference["學年成績單"];
            if (config != null)
            {
                this._select_types.Clear();
                foreach (XmlElement type in config.SelectNodes("Type"))
                {
                    string typeName = type.GetAttribute("Text");
                    if (!this._select_types.ContainsKey(typeName))
                    {
                        this._select_types.Add(typeName, new List<string>());
                    }
                    foreach (XmlElement absence in type.SelectNodes("Absence"))
                    {
                        string absenceName = absence.GetAttribute("Text");
                        if (!this._select_types[typeName].Contains(absenceName))
                        {
                            this._select_types[typeName].Add(absenceName);
                        }
                    }
                }
                if (config.HasAttribute("UseDefault"))
                {
                    this._use_default = bool.Parse(config.GetAttribute("UseDefault"));
                }
                else
                {
                    config.SetAttribute("UseDefault", "True");
                }
                XmlElement customize = (XmlElement)config.SelectSingleNode("CustomizeTemplate");
                if (customize != null)
                {
                    if (!string.IsNullOrEmpty(customize.InnerText))
                    {
                        string templateBase64 = customize.InnerText;
                        this._custom_template = Convert.FromBase64String(templateBase64);
                    }
                }
                else
                {
                    XmlElement newCustomize = config.OwnerDocument.CreateElement("CustomizeTemplate");
                    config.AppendChild(newCustomize);
                }
                XmlElement print = (XmlElement)config.SelectSingleNode("Print");
                if (print != null)
                {
                    this._receive_name = int.Parse(print.GetAttribute("Name"));
                    this._receive_address = int.Parse(print.GetAttribute("Address"));
                    this._resit_sign = print.GetAttribute("ResitSign");
                    this._repeat_sign = print.GetAttribute("RepeatSign");
                    if (!string.IsNullOrEmpty(print.GetAttribute("AllowMoralScoreOver100")))
                    {
                        this._allow_over = bool.Parse(print.GetAttribute("AllowMoralScoreOver100"));
                    }
                }
                else
                {
                    XmlElement newPrint = config.OwnerDocument.CreateElement("Print");
                    newPrint.SetAttribute("Name", "0");
                    newPrint.SetAttribute("Address", "0");
                    newPrint.SetAttribute("ResitSign", "");
                    newPrint.SetAttribute("RepeatSign", "");
                    newPrint.SetAttribute("AllowMoralScoreOver100", "False");
                    this._receive_name = 0;
                    this._receive_address = 0;
                    this._resit_sign = "";
                    this._repeat_sign = "";
                    this._allow_over = false;
                    config.AppendChild(newPrint);
                }
                SmartSchool.Customization.Data.SystemInformation.Preference["學年成績單"] = config;
            }
            else
            {
                config = new XmlDocument().CreateElement("學年成績單");
                SmartSchool.Customization.Data.SystemInformation.Preference["學年成績單"] = config;
            }
        }

        private void LoadSchoolInfo()
        {
            this._school_year = SmartSchool.Customization.Data.SystemInformation.SchoolYear;
        }

        public void Save()
        {
            XmlElement config = SmartSchool.Customization.Data.SystemInformation.Preference["學年成績單"];
            if (config == null)
            {
                config = new XmlDocument().CreateElement("學年成績單");
            }
            config.SetAttribute("UseDefault", this._use_default.ToString());
            XmlElement customize = config.OwnerDocument.CreateElement("CustomizeTemplate");
            XmlElement print = config.OwnerDocument.CreateElement("Print");
            if (this._isUploadNewTemplate)
            {
                customize.InnerText = Convert.ToBase64String(this._custom_template);
                if (config.SelectSingleNode("CustomizeTemplate") == null)
                {
                    config.AppendChild(customize);
                }
                else
                {
                    config.ReplaceChild(customize, config.SelectSingleNode("CustomizeTemplate"));
                }
            }
            print.SetAttribute("Name", this._receive_name.ToString());
            print.SetAttribute("Address", this._receive_address.ToString());
            print.SetAttribute("ResitSign", this._resit_sign);
            print.SetAttribute("RepeatSign", this._repeat_sign);
            print.SetAttribute("AllowMoralScoreOver100", this._allow_over.ToString());
            if (config.SelectSingleNode("Print") == null)
            {
                config.AppendChild(print);
            }
            else
            {
                config.ReplaceChild(print, config.SelectSingleNode("Print"));
            }
            foreach (XmlElement var in config.SelectNodes("Type"))
            {
                config.RemoveChild(var);
            }
            foreach (string type in this._select_types.Keys)
            {
                XmlElement typeElement = config.OwnerDocument.CreateElement("Type");
                typeElement.SetAttribute("Text", "" + type);
                foreach (string absence in this._select_types[type])
                {
                    XmlElement absenceElement = config.OwnerDocument.CreateElement("Absence");
                    absenceElement.SetAttribute("Text", "" + absence);
                    typeElement.AppendChild(absenceElement);
                }
                config.AppendChild(typeElement);
            }
            SmartSchool.Customization.Data.SystemInformation.Preference["學年成績單"] = config;
        }

        public void SetReceiveInfo(int name, int address)
        {
            this._receive_name = name;
            this._receive_address = address;
        }

        public void SetSign(string resit, string repeat)
        {
            this._resit_sign = resit;
            this._repeat_sign = repeat;
        }

        public void SetTypes(Dictionary<string, List<string>> types)
        {
            this._select_types = types;
        }

        public void UploadCustomTemplate(byte[] template)
        {
            this._custom_template = template;
            this._isUploadNewTemplate = true;
        }

        // Properties
        public bool AllowOver
        {
            get
            {
                return this._allow_over;
            }
            set
            {
                this._allow_over = value;
            }
        }

        public byte[] CustomTemplate
        {
            get
            {
                return this._custom_template;
            }
        }

        public byte[] DefaultTemplate
        {
            get
            {
                return Properties.Resources.學年成績單;
            }
        }

        public int ReceiveAddressIndex
        {
            get
            {
                return this._receive_address;
            }
        }

        public int ReceiveNameIndex
        {
            get
            {
                return this._receive_name;
            }
        }

        public string RepeatSign
        {
            get
            {
                return this._repeat_sign;
            }
        }

        public string ResitSign
        {
            get
            {
                return this._resit_sign;
            }
        }

        public decimal SchoolYear
        {
            get
            {
                return this._school_year;
            }
            set
            {
                this._school_year = value;
            }
        }

        public Dictionary<string, List<string>> SelectTypes
        {
            get
            {
                return this._select_types;
            }
        }

        public byte[] Template
        {
            get
            {
                if (this._use_default)
                {
                    return this.DefaultTemplate;
                }
                return this.CustomTemplate;
            }
        }

        public bool UseDefault
        {
            get
            {
                return this._use_default;
            }
            set
            {
                this._use_default = value;
            }
        }
    }

}

