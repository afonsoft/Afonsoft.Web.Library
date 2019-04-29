using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;

namespace Afonsoft.Web.Library
{
    public static class DropDownListExtensions
    {
        /// <summary>
        /// Selecionar um item de um DropDownList
        /// </summary>
        /// <param name="ddl">DropDownList</param>
        /// <param name="value">E o valor que se procura</param>
        public static void SelectByValue(this DropDownList ddl, string value)
        {
            var l = ddl.Items.FindByValue(value);
            var i = ddl.Items.IndexOf(l);
            ddl.SelectedIndex = i;
        }

        /// <summary>
        /// Selecionar um item de um DropDownList
        /// </summary>
        /// <param name="ddl">DropDownList</param>
        /// <param name="text">E o texto que se procura</param>
        public static void SelectByText(this DropDownList ddl, string text)
        {
            var l = ddl.Items.FindByText(text);
            var i = ddl.Items.IndexOf(l);
            ddl.SelectedIndex = i;
        }
    }
}
