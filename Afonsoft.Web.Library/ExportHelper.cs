using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Afonsoft.Web.Library
{
    /// <summary>
    /// Classe para exportação em CSV.
    /// 
    /// </summary>
    public static class ExportHelper
    { /// <summary>
      /// RemoveIvalidCharacters
      /// </summary>
        public static string RemoveIvalidCharacters(string value)
        {
            value = Regex.Replace(value, "[«»\u201C\u201D\u201E\u201F\u2033\u2036]", "");
            value = Regex.Replace(value, "[èëêð]", "e");
            value = Regex.Replace(value, "[ÈËÊ]", "E");
            value = Regex.Replace(value, "[àâä]", "a");
            value = Regex.Replace(value, "[ÀÂÄÅ]", "A");
            value = Regex.Replace(value, "[ÙÛÜ]", "U");
            value = Regex.Replace(value, "[úûüµ]", "u");
            value = Regex.Replace(value, "[òöø]", "o");
            value = Regex.Replace(value, "[ÒÖØ]", "O");
            value = Regex.Replace(value, "[ìîï]", "i");
            value = Regex.Replace(value, "[ÌÎÏ]", "I");
            value = Regex.Replace(value, "[š]", "s");
            value = Regex.Replace(value, "[Š]", "S");
            value = Regex.Replace(value, "[ñ]", "n");
            value = Regex.Replace(value, "[Ñ]", "N");
            value = Regex.Replace(value, "[ÿ]", "y");
            value = Regex.Replace(value, "[Ÿ]", "Y");
            value = Regex.Replace(value, "[ž]", "z");
            value = Regex.Replace(value, "[Ž]", "Z");
            value = Regex.Replace(value, "[Ð]", "D");
            value = Regex.Replace(value, "[œ]", "oe");
            value = Regex.Replace(value, "[Œ]", "Oe");
            value = Regex.Replace(value, "[\"]", "");
            value = Regex.Replace(value, "[;]", "");
            value = Regex.Replace(value, "[\t]", "");
            value = Regex.Replace(value, "[\n]", "");
            value = Regex.Replace(value, "[\r]", "");
            value = Regex.Replace(value, "'", "");
            value = Regex.Replace(value, ";", "");
            value = Regex.Replace(value, "[\u2026]", "...");
            value = Regex.Replace(value, Environment.NewLine, "");
            return value;
        }

        private static string TratarStr(string str)
        {
            return TratarStr(RemoveIvalidCharacters(str), false);
        }
        private static string TratarStr(string str, bool exibirTexto)
        {
            if (exibirTexto)
                return "=\"" + RemoveIvalidCharacters(str) + "\"";
            return "\"" + RemoveIvalidCharacters(str) + "\"";
        }
        /// <summary>
        /// Exportar um DataTable para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, DataTable dt)
        {
            Export(fileName, dt, false);
        }
        /// <summary>
        /// Exportar um DataTable para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, DataTable dt, bool removeSpaceTitle)
        {
            Export(fileName, dt, removeSpaceTitle, Encoding.GetEncoding("ISO-8859-1"));
        }
        /// <summary>
        /// Exportar um DataTable para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, DataTable dt, bool removeSpaceTitle, Encoding encoding)
        {
            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.ClearContent();
            HttpContext.Current.Response.ContentEncoding = encoding;
            //HttpContext.Current.Response.BinaryWrite(encoding.GetPreamble()); //Clean GZip Encode Page
            if (HttpContext.Current.Response.Filter.GetType() == typeof(System.IO.Compression.GZipStream))
                HttpContext.Current.Response.Filter = null; //Clean GZip Encode Page
            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            HttpContext.Current.Response.AddHeader("content-disposition", $"attachment; filename={fileName}");
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";

            StringBuilder sb = new StringBuilder();
            //Add Header    
            for (int count = 0; count < dt.Columns.Count; count++)
            {
                if (dt.Columns[count] != null)
                    sb.Append(removeSpaceTitle
                        ? TratarStr(dt.Columns[count].ColumnName).Replace("__", ". ").Replace("_", " ")
                        : TratarStr(dt.Columns[count].ColumnName));
                if (count < dt.Columns.Count - 1)
                {
                    sb.Append(";");
                }
            }
            HttpContext.Current.Response.Write(sb + "\n");
            HttpContext.Current.Response.Flush();
            //Append Data

            for (int count = 0; count < dt.Rows.Count; count++)
            {
                sb = new StringBuilder();
                for (int col = 0; col <= dt.Columns.Count - 1; col++)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[count][col].ToString()))
                        sb.Append(TratarStr(dt.Rows[count][col].ToString()));
                    sb.Append(";");
                }

                HttpContext.Current.Response.Write(sb + "\n");
                HttpContext.Current.Response.Flush();
            }
            //dr.Dispose();

            HttpContext.Current.Response.End();
        }

        /// <summary>
        /// Exportar um objeto para o excell somente para exibição na formtação somente texto em todos os campos.
        /// </summary>
        public static void ExportOnlyViewInExcel<T>(string fileName, IEnumerable<T> obj)
        {
            ExportOnlyViewInExcel(fileName, obj, false);
        }
        /// <summary>
        /// Exportar um objeto para o excell somente para exibição na formtação somente texto em todos os campos.
        /// </summary>
        public static void ExportOnlyViewInExcel<T>(string fileName, IEnumerable<T> obj, bool removeSpaceTitle)
        {
            ExportOnlyViewInExcel(fileName, obj, removeSpaceTitle, Encoding.GetEncoding("ISO-8859-1"));
        }
        /// <summary>
        /// Exportar um objeto para o excell somente para exibição na formtação somente texto em todos os campos.
        /// </summary>
        public static void ExportOnlyViewInExcel<T>(string fileName, IEnumerable<T> obj, bool removeSpaceTitle, Encoding encoding)
        {
            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.ClearContent();
            HttpContext.Current.Response.ContentEncoding = encoding;
            //HttpContext.Current.Response.BinaryWrite(encoding.GetPreamble()); //Clean GZip Encode Page
            if (HttpContext.Current.Response.Filter.GetType() == typeof(System.IO.Compression.GZipStream))
                HttpContext.Current.Response.Filter = null; //Clean GZip Encode Page
            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            HttpContext.Current.Response.AddHeader("content-disposition", $"attachment; filename={fileName}");
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";

            if (obj == null)
                return;

            var enumerable = obj as T[] ?? obj.ToArray();
            Type type = enumerable.ToArray()[0].GetType();

            StringBuilder sb = new StringBuilder();

            //Add Header  
            for (int count = 0; count < type.GetProperties().Length; count++)
            {
                System.Reflection.PropertyInfo property = type.GetProperties()[count];
                if (!string.IsNullOrEmpty(property.Name))
                {
                    sb.Append(removeSpaceTitle
                        ? TratarStr(property.Name).Replace("__", ". ").Replace("_", " ")
                        : TratarStr(property.Name, true));

                    if (count < type.GetProperties().Length - 1)
                    {
                        sb.Append(";");
                    }
                }
            }

            HttpContext.Current.Response.Write(sb + "\n");
            HttpContext.Current.Response.Flush();

            //Append Data
            foreach (T t in enumerable)
            {
                //Varrendo item por item do objeto
                type = t.GetType();

                sb = new StringBuilder();

                for (int count = 0; count < type.GetProperties().Length; count++)
                {
                    System.Reflection.PropertyInfo property = type.GetProperties()[count];

                    if (!string.IsNullOrEmpty(property.Name))
                    {
                        object value = type.InvokeMember(property.Name, System.Reflection.BindingFlags.GetProperty, null, t, new object[0]);

                        if (value != null && !string.IsNullOrEmpty(value.ToString()))
                        {
                            value = value.ToString().Trim();
                            sb.Append(TratarStr(value.ToString(), true));
                        }
                        sb.Append(";");
                    }

                }

                HttpContext.Current.Response.Write(sb + "\n");
                HttpContext.Current.Response.Flush();
            }

            HttpContext.Current.Response.End();
        }
        /// <summary>
        /// Exportar um IEnumerable para o excel (.csv)
        /// </summary>
        public static void Export<T>(string fileName, IEnumerable<T> obj)
        {
            Export(fileName, obj, false);
        }

        /// <summary>
        /// Exportar um IEnumerable para o excel (.csv)
        /// </summary>
        public static void Export<T>(string fileName, IEnumerable<T> obj, bool removeSpaceTitle)
        {
            Export(fileName, obj, removeSpaceTitle, Encoding.GetEncoding("ISO-8859-1"));
        }

        /// <summary>
        /// Exportar um IEnumerable para o excel
        /// </summary>
        public static void Export<T>(string fileName, IEnumerable<T> obj, bool removeSpaceTitle, Encoding encoding)
        {
            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.ClearContent();
            HttpContext.Current.Response.ContentEncoding = encoding;
            //HttpContext.Current.Response.BinaryWrite(encoding.GetPreamble()); //Clean GZip Encode Page
            if (HttpContext.Current.Response.Filter.GetType() == typeof(System.IO.Compression.GZipStream))
                HttpContext.Current.Response.Filter = null; //Clean GZip Encode Page
            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            HttpContext.Current.Response.AddHeader("content-disposition", $"attachment; filename={fileName}");
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";

            if (obj == null)
                return;

            var enumerable = obj as T[] ?? obj.ToArray();
            Type type = enumerable.ToArray()[0].GetType();

            StringBuilder sb = new StringBuilder();

            //Add Header  
            for (int count = 0; count < type.GetProperties().Length; count++)
            {
                System.Reflection.PropertyInfo property = type.GetProperties()[count];
                if (!string.IsNullOrEmpty(property.Name))
                {
                    sb.Append(removeSpaceTitle
                        ? TratarStr(property.Name).Replace("__", ". ").Replace("_", " ")
                        : TratarStr(property.Name));

                    if (count < type.GetProperties().Length - 1)
                    {
                        sb.Append(";");
                    }
                }
            }

            HttpContext.Current.Response.Write(sb + "\n");
            HttpContext.Current.Response.Flush();

            //Append Data
            foreach (T t in enumerable)
            {
                //Varrendo item por item do objeto
                type = t.GetType();

                sb = new StringBuilder();

                for (int count = 0; count < type.GetProperties().Length; count++)
                {
                    System.Reflection.PropertyInfo property = type.GetProperties()[count];

                    if (!string.IsNullOrEmpty(property.Name))
                    {
                        object value = type.InvokeMember(property.Name, System.Reflection.BindingFlags.GetProperty, null, t, new object[0]);

                        if (value != null && !string.IsNullOrEmpty(value.ToString()))
                        {
                            value = value.ToString().Trim();
                            sb.Append(TratarStr(value.ToString()));
                        }
                        sb.Append(";");
                    }

                }

                HttpContext.Current.Response.Write(sb + "\n");
                HttpContext.Current.Response.Flush();
            }

            HttpContext.Current.Response.End();
        }


        /// <summary>
        /// Exportar um DataView para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, DataView dv)
        {
            Export(fileName, dv, false);
        }

        /// <summary>
        /// Exportar um DataView para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, DataView dv, bool removeSpaceTitle)
        {
            Export(fileName, dv, removeSpaceTitle, Encoding.GetEncoding("ISO-8859-1"));
        }

        /// <summary>
        /// Exportar um DataView para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, DataView dv, bool removeSpaceTitle, Encoding encoding)
        {
            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.ClearContent();
            HttpContext.Current.Response.ContentEncoding = encoding;
            //HttpContext.Current.Response.BinaryWrite(encoding.GetPreamble()); //Clean GZip Encode Page
            if (HttpContext.Current.Response.Filter.GetType() == typeof(System.IO.Compression.GZipStream))
                HttpContext.Current.Response.Filter = null; //Clean GZip Encode Page
            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            HttpContext.Current.Response.AddHeader("content-disposition", $"attachment; filename={fileName}");
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";

            var dt = dv.ToTable();

            StringBuilder sb = new StringBuilder();
            //Add Header    
            for (int count = 0; count < dt.Columns.Count; count++)
            {
                if (dt.Columns[count] != null)
                    sb.Append(removeSpaceTitle
                        ? TratarStr(dt.Columns[count].ColumnName).Replace("__", ". ").Replace("_", " ")
                        : TratarStr(dt.Columns[count].ColumnName));
                if (count < dt.Columns.Count - 1)
                {
                    sb.Append(";");
                }
            }
            HttpContext.Current.Response.Write(sb + "\n");
            HttpContext.Current.Response.Flush();
            //Append Data

            for (int count = 0; count < dt.Rows.Count; count++)
            {
                sb = new StringBuilder();
                for (int col = 0; col <= dt.Columns.Count - 1; col++)
                {
                    if (!string.IsNullOrEmpty(dt.Rows[count][col].ToString()))
                        sb.Append(TratarStr(dt.Rows[count][col].ToString()));
                    sb.Append(";");
                }

                HttpContext.Current.Response.Write(sb + "\n");
                HttpContext.Current.Response.Flush();
            }
            //dr.Dispose();

            HttpContext.Current.Response.End();
        }

        /// <summary>
        /// Exportar um IDataReader para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, IDataReader dr)
        {
            Export(fileName, dr, false);
        }
        /// <summary>
        /// Exportar um IDataReader para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, IDataReader dr, bool removeSpaceTitle)
        {
            Export(fileName, dr, removeSpaceTitle, Encoding.GetEncoding("ISO-8859-1"));
        }

        /// <summary>
        /// Exportar um IDataReader para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, IDataReader dr, bool removeSpaceTitle, Encoding encoding)
        {

            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.ClearContent();
            HttpContext.Current.Response.ContentEncoding = encoding;
            //HttpContext.Current.Response.BinaryWrite(encoding.GetPreamble()); //Clean GZip Encode Page
            if (HttpContext.Current.Response.Filter.GetType() == typeof(System.IO.Compression.GZipStream))
                HttpContext.Current.Response.Filter = null; //Clean GZip Encode Page
            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            HttpContext.Current.Response.AddHeader("content-disposition", $"attachment; filename={fileName}");
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";

            StringBuilder sb = new StringBuilder();
            //Add Header          

            for (int count = 0; count < dr.FieldCount; count++)
            {
                if (dr.GetName(count) != null)
                    sb.Append(removeSpaceTitle
                        ? TratarStr(dr.GetName(count)).Replace("__", ". ").Replace("_", " ")
                        : TratarStr(dr.GetName(count)));
                if (count < dr.FieldCount - 1)
                {
                    sb.Append(";");
                }
            }
            HttpContext.Current.Response.Write(sb + "\n");
            HttpContext.Current.Response.Flush();
            //Append Data

            while (dr.Read())
            {
                sb = new StringBuilder();

                for (int col = 0; col < dr.FieldCount - 1; col++)
                {
                    if (!dr.IsDBNull(col))
                        sb.Append(TratarStr(dr.GetValue(col).ToString()));

                    sb.Append(";");
                }
                if (!dr.IsDBNull(dr.FieldCount - 1))
                    sb.Append(TratarStr(dr.GetValue(dr.FieldCount - 1).ToString()));
                HttpContext.Current.Response.Write(sb + "\n");
                HttpContext.Current.Response.Flush();
            }
            dr.Dispose();

            HttpContext.Current.Response.End();
        }

        /// <summary>
        /// Exportar um DataSet para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, DataSet ds)
        {
            Export(fileName, ds, false);
        }

        /// <summary>
        /// Exportar um DataSet para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, DataSet ds, bool removeSpaceTitle)
        {
            if (ds.Tables.Count > 0)
                Export(fileName, ds.Tables[0], removeSpaceTitle, Encoding.GetEncoding("ISO-8859-1"));
        }

        /// <summary>
        /// Exportar um DataSet para o excel (.csv)
        /// </summary>
        public static void Export(string fileName, DataSet ds, bool removeSpaceTitle, Encoding encoding)
        {
            if (ds.Tables.Count > 0)
                Export(fileName, ds.Tables[0], removeSpaceTitle, encoding);
        }

        /// <summary>
        /// Exportar um GridView para o excel (XLS 65000 linhas)
        /// </summary>
        public static void Export(string fileName, GridView gv)
        {
            Export(fileName, gv, Encoding.GetEncoding("ISO-8859-1"));
        }

        /// <summary>
        /// Exportar um GridView para o excel (XLS 65000 linhas)
        /// </summary>
        public static void Export(string fileName, GridView gv, Encoding encoding)
        {
            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.ClearContent();
            HttpContext.Current.Response.ContentEncoding = encoding;
            //HttpContext.Current.Response.BinaryWrite(encoding.GetPreamble()); //Clean GZip Encode Page
            if (HttpContext.Current.Response.Filter.GetType() == typeof(System.IO.Compression.GZipStream))
                HttpContext.Current.Response.Filter = null; //Clean GZip Encode Page
            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            HttpContext.Current.Response.AddHeader("content-disposition", $"attachment; filename={fileName}");
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";

            gv.AllowPaging = false;
            gv.AutoGenerateSelectButton = false;
            PrepareControlForExport(gv);

            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    gv.RenderControl(htw);
                    //  render the htmlwriter into the response
                    String html = sw.ToString().Replace(Environment.NewLine, "").Replace("&nbsp;", "");
                    HttpContext.Current.Response.Write(html);
                    HttpContext.Current.Response.Flush();
                }
            }
            HttpContext.Current.Response.End();
        }

        /// <summary>
        /// Exportar um Repeater para o excel (XLS 65000 linhas)
        /// </summary>
        public static void Export(string fileName, Repeater rep)
        {
            Export(fileName, rep, Encoding.GetEncoding("ISO-8859-1"));
        }

        /// <summary>
        /// Exportar um Repeater para o excel (XLS 65000 linhas)
        /// </summary>
        public static void Export(string fileName, Repeater rep, Encoding encoding)
        {
            HttpContext.Current.Response.Buffer = true;
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ClearHeaders();
            HttpContext.Current.Response.ClearContent();
            HttpContext.Current.Response.ContentEncoding = encoding;
            //HttpContext.Current.Response.BinaryWrite(encoding.GetPreamble()); //Clean GZip Encode Page
            if (HttpContext.Current.Response.Filter.GetType() == typeof(System.IO.Compression.GZipStream))
                HttpContext.Current.Response.Filter = null; //Clean GZip Encode Page
            HttpContext.Current.Response.Cache.SetCacheability(HttpCacheability.NoCache);
            HttpContext.Current.Response.AddHeader("content-disposition", $"attachment; filename={fileName}");
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";

            PrepareControlForExport(rep);

            using (StringWriter sw = new StringWriter())
            {
                using (HtmlTextWriter htw = new HtmlTextWriter(sw))
                {
                    rep.RenderControl(htw);
                    //  render the htmlwriter into the response
                    String html = sw.ToString().Replace(Environment.NewLine, "").Replace("&nbsp;", "");
                    HttpContext.Current.Response.Write(html);
                    HttpContext.Current.Response.Flush();
                }
            }
            HttpContext.Current.Response.End();
        }

        /// <summary>
        /// Replace any of the contained controls with literals
        /// </summary>
        /// <param name="control"></param>
        private static void PrepareControlForExport(Control control)
        {
            for (int i = 0; i < control.Controls.Count; i++)
            {
                Control current = control.Controls[i];
                if (current is LinkButton button)
                {
                    control.Controls.Remove(button);
                    control.Controls.AddAt(i, new LiteralControl(button.Text));
                }
                else if (current is ImageButton imageButton)
                {
                    control.Controls.Remove(imageButton);
                    control.Controls.AddAt(i, new LiteralControl(imageButton.AlternateText));
                }
                else if (current is HyperLink link)
                {
                    control.Controls.Remove(link);
                    control.Controls.AddAt(i, new LiteralControl(link.Text));
                }
                else if (current is DropDownList list)
                {
                    control.Controls.Remove(list);
                    control.Controls.AddAt(i, new LiteralControl(list.SelectedItem.Text));
                }
                else if (current is CheckBox box)
                {
                    control.Controls.Remove(box);
                    control.Controls.AddAt(i, new LiteralControl(box.Checked ? "Sim" : "Não"));
                }
                else if (current is HiddenField)
                {
                    control.Controls.Remove(current);
                }

                if (current.HasControls())
                {
                    PrepareControlForExport(current);
                }
            }
        }
    }
}