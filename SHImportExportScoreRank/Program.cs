using Campus.DocumentValidator;
using FISCA.Permission;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using SHImportExportScoreRank.Permissions;
using System;

namespace SHImportExportScoreRank
{
    public class Program
    {
        private static bool _initialized;

        [FISCA.MainMethod()]
        public static void Main()
        {
            if (_initialized)
                return;

            _initialized = true;

            RegisterPermissions();
            RegisterImportFeature();
            RegisterExportFeature();
            RegisterDetailContent();
            RegisterValidators();
        }

        private static void RegisterPermissions()
        {
            Catalog catalog =
                RoleAclSource.Instance["學生"]["學年分項排名"];

            catalog.Add(
                new RibbonFeature(
                    FeatureCode.ImportSchoolYearEntryRank,
                    "匯入學年分項排名"));

            catalog.Add(
                new RibbonFeature(
                    FeatureCode.ExportSchoolYearEntryRank,
                    "匯出學年分項排名"));

            catalog.Add(
                new DetailItemFeature(
                    FeatureCode.StudentSchoolYearEntryRankDetail,
                    "排名資料"));
        }

        private static void RegisterImportFeature()
        {
            RibbonBarButton importRoot =
                MotherForm.RibbonBarItems["學生", "資料統計"]["匯入"];

            importRoot["成績相關匯入"]["匯入學年分項排名"].Enable =
                UserAcl.Current[
                    FeatureCode.ImportSchoolYearEntryRank
                ].Executable;

            importRoot["成績相關匯入"]["匯入學年分項排名"].Click += delegate
            {
                if (!UserAcl.Current[
                        FeatureCode.ImportSchoolYearEntryRank
                    ].Executable)
                {
                    MsgBox.Show("您沒有「匯入學年分項排名」權限。");
                    return;
                }

                try
                {
                    ImportExport.ImportSchoolYearEntryRank importer =
                        new ImportExport.ImportSchoolYearEntryRank();

                    importer.Execute();
                }
                catch (Exception ex)
                {
                    FISCA.LogAgent.ApplicationLog.Log(
                        "匯入學年分項排名",
                        "啟動失敗",
                        ex.ToString());

                    MsgBox.Show(
                        "啟動匯入學年分項排名功能失敗：" +
                        ex.Message);
                }
            };
        }

        private static void RegisterExportFeature()
        {
            RibbonBarButton exportRoot =
                MotherForm.RibbonBarItems["學生", "資料統計"]["匯出"];

            exportRoot["成績相關匯出"]["匯出學年分項排名"].Enable =
                UserAcl.Current[
                    FeatureCode.ExportSchoolYearEntryRank
                ].Executable;

            exportRoot["成績相關匯出"]["匯出學年分項排名"].Click += delegate
            {
                if (!UserAcl.Current[
                        FeatureCode.ExportSchoolYearEntryRank
                    ].Executable)
                {
                    MsgBox.Show("您沒有「匯出學年分項排名」權限。");
                    return;
                }

                try
                {
                    SmartSchool.API.PlugIn.Export.Exporter exporter =
                        new ImportExport.ExportSchoolYearEntryRank();

                    ImportExport.ExportStudentV2 wizard =
                        new ImportExport.ExportStudentV2(
                            exporter.Text,
                            exporter.Image);

                    exporter.InitializeExport(wizard);
                    wizard.ShowDialog();
                }
                catch (Exception ex)
                {
                    FISCA.LogAgent.ApplicationLog.Log(
                        "學年分項排名",
                        "匯出啟動失敗",
                        ex.ToString());

                    MsgBox.Show(
                        "啟動匯出學年分項排名功能失敗：" +
                        ex.Message);
                }
            };
        }

        private static void RegisterDetailContent()
        {
            if (!UserAcl.Current[
                    FeatureCode.StudentSchoolYearEntryRankDetail
                ].Viewable)
            {
                return;
            }

            K12.Presentation.NLDPanels.Student.AddDetailBulider(
                new FISCA.Presentation.DetailBulider<
                    DetailContent.UCStudentRank>());
        }

        private static void RegisterValidators()
        {
            FactoryProvider.RowFactory.Add(
                new ValidationRule
                    .SchoolYearEntryRankRowValidatorFactory());
        }
    }
}
