using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace RTSafe.HiddenTroubleTreatm.BusinessModules.HiddenTroubleTreatmModules
{
    public class ExcelToModel
    {
        /// <summary>
        /// 获取excel数据集合
        /// </summary>
        /// <typeparam name="T">对象</typeparam>
        /// <param name="ws">excel数据流</param>
        /// <param name="xmlNodeName">节点对象</param>
        /// <returns></returns>
        public List<T> ReadTable<T>(Worksheet ws, string xmlNodeName)
        {

            Cells all = ws.Cells;
            DataSet ds = new DataSet();
            ds.ReadXml(System.Web.HttpContext.Current.Request.MapPath("~/ImportFiles/InputRules.xml"));
            DataTable dt = ds.Tables[xmlNodeName];
            // 数据区
            int idx = -1;

            List<T> list = new List<T>();
            for (int i = 1; i < all.Rows.Count; i++)
            {
                T model = Activator.CreateInstance<T>();
                Dictionary<string, object> dic = new Dictionary<string, object>();
                for (int J = 0; J < ds.Tables[xmlNodeName].Rows.Count; J++)
                {
                    #region 获取excel表格中的值
                    int ExcelColumn = int.Parse(ds.Tables[xmlNodeName].Rows[J]["ExcelColumn"].ToString());
                    string OracleType = ds.Tables[xmlNodeName].Rows[J]["OracleType"].ToString();
                    string OracleColumn = ds.Tables[xmlNodeName].Rows[J]["OracleColumn"].ToString();

                    string ParameterName = ds.Tables[xmlNodeName].Rows[J]["ParameterName"].ToString();
                    string MethodName = ds.Tables[xmlNodeName].Rows[J]["MethodName"].ToString();
                    string strValue = string.Empty;
                    strValue = ExcelColumn > -1 ? all[i, ExcelColumn].StringValue : "";
                    if (!Verification(ExcelColumn, OracleType, strValue))
                        break;
                    if (!string.IsNullOrEmpty(ParameterName.ToString()))
                    {
                        string ParameterValue = all[i, int.Parse(ParameterName)].StringValue;
                        var factory = CustomParmeter_Factory.Cresate_CustomParmeter_Factory(MethodName, ParameterValue);
                        strValue = factory.GetCustomParmeter();
                        if (string.IsNullOrEmpty(strValue))
                            break;
                    }
                    dic.Add(OracleColumn, ObjConvert(strValue, OracleType));
                    #endregion
                }
                PropertyInfo[] modelPro = model.GetType().GetProperties();
                if (modelPro.Length > 0 && dic.Count() > 0)
                {
                    for (int z = 0; z < modelPro.Length; z++)
                    {
                        if (dic.ContainsKey(modelPro[z].Name))
                        {
                            modelPro[z].SetValue(model, dic[modelPro[z].Name], null);
                        }

                    }
                    list.Add(model);
                }
            }
            return list;
        }
        public object ObjConvert(string strValue, string OracleType)
        {
            object obj;
            switch (OracleType)
            {
                case "PK":
                    obj = Guid.NewGuid();
                    break;
                case "datetime":
                    obj = Convert.ToDateTime(strValue);
                    break;
                case "decimal":
                    obj = Convert.ToDecimal(strValue);
                    break;
                case "int":
                    obj = Convert.ToDecimal(strValue);
                    break;
                case "Guid":
                    obj = new Guid(strValue);
                    break;
                default:
                    obj = strValue;
                    break;
            }
            return obj;
        }
        public static bool Verification(int ExcelColumn, string OracleType, string strValue)
        {
            DateTime dtDate;
            if (ExcelColumn < 0)
                return true;
            if (string.IsNullOrEmpty(OracleType) || string.IsNullOrEmpty(strValue))
                return false;
            if (OracleType == "datetime" && !DateTime.TryParse(strValue, out dtDate))
                return false;
            return true;
        }
    }
}
