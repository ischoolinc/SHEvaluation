namespace SchoolYearScoreReport
{
    using SmartSchool.AccessControl;
    using SmartSchool.Customization.PlugIn;
    using System;

    internal class SecureButtonAdapter : ButtonAdapter, IFeature
    {
        private string _feature_code;

        public SecureButtonAdapter(string featureCode)
        {
            this._feature_code = featureCode;
        }

        public string FeatureCode
        {
            get
            {
                return this._feature_code;
            }
        }
    }
}

