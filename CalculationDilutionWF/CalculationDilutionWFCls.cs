using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using XmlService;

namespace CalculationDilutionWF
{

    [ComVisible(true)]
    [ProgId("CalculationDilutionWF.CalculationDilutionWFCls")]
    public class CalculationDilutionWFCls : IWorkflowExtension
    {

        INautilusServiceProvider sp;
        private const string FinalResultName = "Final Result";
        public void Execute(ref LSExtensionParameters Parameters)
        {
            DataLayer dal = null;
            try
            {
                //to be removed !!!!!
                Logger.WriteLogFile("entered", false);
                //
                sp = Parameters["SERVICE_PROVIDER"];

                var rs = Parameters["RECORDS"];

                var resultId = rs.Fields["RESULT_ID"].Value;

                var ntlsCon = Utils.GetNtlsCon(sp);

                Utils.CreateConstring(ntlsCon);

                dal = new DataLayer();
                dal.Connect();

                //Get Specified result
                Result result = dal.GetResultById(long.Parse(resultId.ToString()));

                //Get Parent test
                var test = result.Test;

                //Get all results
                var results = test.Results.ToList();

                //Get final result
                var finalResult = results.Where(r => r.Name == FinalResultName).FirstOrDefault();

                //cancek ���� �� �3 ������� �������� �� ������� ����� 
                //???Please add this condition : "or (result_status="C" and original_result is NULL)"
                //var notRejecteds = results.Where(r => (r.Status != "X" && (r.Status == "C" && r.ORIGINAL_RESULT == null)) && r.ResultId != finalResult.ResultId).ToList();
                var notRejecteds =
                    results.Where(r => (r.FormattedResult != "Not Entered" && r.ResultId != finalResult.ResultId))
                        .ToList();


                //reject �� �� ����� ��� �� ???Please add this condition : "or (result_status="C" and original_result is NULL)"
                //reject �� ������ ����� ��� ???Please add this condition : "or (result_status="C" and original_result is NULL)"

                var result2 =
                    notRejecteds.Where(
                        r => r.ResultTemplate.Name == "����� ����� ���" || r.ResultTemplate.Name == "Dilution 2")
                        .FirstOrDefault();

                if (finalResult == null || result2 == null || notRejecteds.Count() != 2)
                {
                    //???Sefi will think what happens in this case
                    Logger.WriteLogFile("�� ������� ������ ������ ����.", false);
                    MessageBox.Show("�� ������� ������ ������ ����.", "Nautilus");
                    return;
                }
                //Sum rejected results
                double sum = notRejecteds.Sum(r => (double)r.RAW_NUMERIC_RESULT);

                //���� �������� ����� ������ �� �� ������� ��� ���� ����
                double minDF = (double)notRejecteds.Min(r => r.DilutionFactor);

                //������ ������ ��� ������ �������
                var resultValue = sum / 1.1 * Math.Pow(10, Convert.ToDouble(-minDF));
                var roundUpValue = Math.Ceiling(resultValue);
                //Set value to final result
                var reXml = new ResultEntryXmlHandler(sp);
                reXml.CreateResultEntryXml(test.TEST_ID, finalResult.ResultId, roundUpValue.ToString());
                var res = reXml.ProcssXml();
                Logger.WriteLogFile("Calculation DilutionWF success", false);
            }
            catch (Exception ex)
            {
                Logger.WriteLogFile(ex);
                MessageBox.Show("Error " + ex.Message);

            }
            finally
            {
                if (dal != null) dal.Close();
            }
        }


    }
}
