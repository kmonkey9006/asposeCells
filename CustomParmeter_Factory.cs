using RTSafe.HiddenTroubleTreatm.BusinessModules.HiddenTroubleTreatmModules.Domain;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RTSafe.HiddenTroubleTreatm.BusinessModules.HiddenTroubleTreatmModules
{
    public class CustomParmeter_Factory
    {

        public static ICustomParmeter Cresate_CustomParmeter_Factory(string MethodName, string ParameterName)
        {
            ICustomParmeter newService = null;
            switch (MethodName)
            {
                case "SetLine":
                    newService = new SetLine(ParameterName);
                    break;
                case "SetSegment":
                    newService = new SetSegment(ParameterName);
                    break;
                case "SetSite":
                    newService = new SetSite(ParameterName);
                    break;
            }
            return newService;
        }
    }
    public interface ICustomParmeter
    {
        string GetCustomParmeter();
    }

    public class SetSegment : ICustomParmeter
    {

        private string ParameterName;


        public SetSegment(string _ParameterName)
        {
            ParameterName = _ParameterName;
        }
        public string GetCustomParmeter()
        {
            CustomParmeterCommonService c = new CustomParmeterCommonService();
            return c.GetSegmentPatmet(ParameterName).ToString();


        }
    }

    public class SetLine : ICustomParmeter
    {

        private static string ParameterName;

        public SetLine(string _ParameterName)
        {
            ParameterName = _ParameterName;
        }
        public string GetCustomParmeter()
        {
            CustomParmeterCommonService c = new CustomParmeterCommonService();
            return c.GetLineParmet(ParameterName).ToString();
        }
    }
    public class SetSite : ICustomParmeter
    {
     

        private string ParameterName;
        public SetSite(string _ParameterName)
        {
            ParameterName = _ParameterName;
        }
        public string GetCustomParmeter()
        {
            CustomParmeterCommonService c = new CustomParmeterCommonService();
            return c.GetSitePatmet(ParameterName).ToString();
        }
    }

}
