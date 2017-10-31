using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;

namespace Tools
{
    class Data2Conversion
    {

        //将xml转为Datable
        public static DataTable ConvertXml2DataTable(string xmlData)
        {
            StringReader stream = null;
            XmlTextReader reader = null;
            try
            {
                DataSet xmlDS = new DataSet();
                stream = new StringReader(xmlData);
                reader = new XmlTextReader(stream);
                xmlDS.ReadXml(reader);

                return xmlDS.Tables[0];
            }
            catch (Exception ex)
            {
                string strTest = ex.Message;
                return null;
            }
            finally
            {
                if (reader != null) reader.Close();
            }
        }
        public static DataSet ConvertXml2DataSet(string xmlData)
        {
            DataSet ds = new DataSet();
            TextReader tr = new StringReader(xmlData);
            ds.ReadXml(tr);
            return ds;
        }
        /// <summary>
        /// Array转DataTable
        /// </summary>
        /// <param name="ColumnName">列名</param>
        /// <param name="Array">传入数组</param>
        /// <returns>DataTable</returns>
        public static DataTable Array2DataTable(string ColumnName, string[] Array)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(ColumnName, typeof(string));

            for (int i = 0; i < Array.Length; i++)
            {
                DataRow dr = dt.NewRow();
                dr[ColumnName] = Array[i].ToString();
                dt.Rows.Add(dr);
            }

            return dt;
        }

        public static DataTable List2DataTable<T>(IEnumerable<T> collection)
        {
            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p => new DataColumn(p.Name, p.PropertyType)).ToArray());
            if (collection.Count() > 0)
            {
                for (int i = 0; i < collection.Count(); i++)
                {
                    ArrayList tempList = new ArrayList();
                    foreach (PropertyInfo pi in props)
                    {
                        object obj = pi.GetValue(collection.ElementAt(i), null);
                        tempList.Add(obj);
                    }
                    object[] array = tempList.ToArray();
                    dt.LoadDataRow(array, true);
                }
            }
            return dt;
        }
        
    }
}
