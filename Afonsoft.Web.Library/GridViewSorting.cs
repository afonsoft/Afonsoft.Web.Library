using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Web;
using System.Web.UI.WebControls;

namespace Afonsoft.Web.Library
{
    /// <summary>
    /// Classe para criar a ordernação do GridView
    /// </summary>
    public class GridViewSorting : System.Web.UI.Control
    {

        #region getSortDirection
        private static string GetSortDirection(string clientId, SortDirection sortDirection)
        {
            string r;
            if (HttpContext.Current.Session["sortDirection_" + clientId] == null)
                HttpContext.Current.Session["sortDirection_" + clientId] = sortDirection;
            sortDirection = (SortDirection)HttpContext.Current.Session["sortDirection_" + clientId] == SortDirection.Ascending ? SortDirection.Descending : SortDirection.Ascending;

            if (sortDirection == SortDirection.Ascending)
                r = "ASC";
            else if (sortDirection == SortDirection.Descending)
                r = "DESC";
            else
                r = "ASC";
            HttpContext.Current.Session["sortDirection_" + clientId] = sortDirection;
            return r;
        }
        #endregion

        #region getSortDataSet
        /// <summary>
        /// Metodo para efetuar um sort em um DataSet
        /// </summary>
        public static DataSet GetSortDataSet(string clientId, DataSet ds, string sortExpression, SortDirection sortDirection)
        {
            if (ds != null && ds.Tables.Count > 0)
            {
                string tableName = string.IsNullOrEmpty(clientId) ? ds.Tables[0].TableName : clientId;
                DataView dv = ds.Tables[0].DefaultView;
                dv.Sort = "[" + sortExpression + "] " + GetSortDirection(tableName, sortDirection);
                DataTable tb = dv.ToTable();
                DataSet dsR = new DataSet();
                dsR.Tables.Add(tb);
                return dsR;
            }
            return null;
        }

        /// <summary>
        /// Metodo para efetuar um sort em um DataSet
        /// </summary>
        public static DataSet GetSortDataSet(DataSet ds, string sortExpression, SortDirection sortDirection)
        {
            return GetSortDataSet(sortExpression, ds, sortExpression, sortDirection);
        }

        #endregion

        #region getSortDataTable

        /// <summary>
        /// Metodo para efetuar um sort em um DataTable
        /// </summary>
        public static DataTable GetSortDataTable(string clientId, DataTable dt, string sortExpression, SortDirection sortDirection)
        {
            if (dt != null)
            {
                string tableName = string.IsNullOrEmpty(clientId) ? dt.TableName : clientId;
                DataView dv = dt.DefaultView;
                dv.Sort = "[" + sortExpression + "] " + GetSortDirection(tableName, sortDirection);
                DataTable tb = dv.ToTable();
                return tb;
            }
            return null;
        }

        /// <summary>
        /// Metodo para efetuar um sort em um DataTable
        /// </summary>
        public static DataTable GetSortDataTable(DataTable dt, string sortExpression, SortDirection sortDirection)
        {
            return GetSortDataTable(sortExpression, dt, sortExpression, sortDirection);
        }

        #endregion

        #region getSortDataView
        /// <summary>
        /// Metodo para efetuar um sort em um DataView
        /// </summary>
        public static DataView GetSortDataView(string clientId, DataView dv, string sortExpression, SortDirection sortDirection)
        {
            if (dv != null)
            {
                string tableName = string.IsNullOrEmpty(clientId) ? sortExpression : clientId;
                dv.Sort = "[" + sortExpression + "] " + GetSortDirection(tableName, sortDirection);
                return dv;
            }
            return null;
        }

        /// <summary>
        /// Metodo para efetuar um sort em um DataView
        /// </summary>
        public static DataView GetSortDataView(DataView dv, string sortExpression, SortDirection sortDirection)
        {
            return GetSortDataView(sortExpression, dv, sortExpression, sortDirection);
        }
        #endregion

        #region getSortObject
        /// <summary>
        /// Metodo para efetuar um sort em um objeto do tipo IEnumerable
        /// </summary>
        public static IEnumerable<T> GetSortObject<T>(IEnumerable<T> obj, string sortExpression, SortDirection sortDirection)
        {
            var param = Expression.Parameter(typeof(T), sortExpression);
            var sortExpress = Expression.Lambda<Func<T, object>>(Expression.Convert(Expression.Property(param, sortExpression), typeof(object)), param);

            SortDirection newsortDirection;
            if (HttpContext.Current.Session[sortExpression + obj.GetType()] != null)
                newsortDirection = (SortDirection)HttpContext.Current.Session[sortExpression + obj.GetType()];
            else
                newsortDirection = sortDirection;

            if (newsortDirection == sortDirection && sortDirection == SortDirection.Ascending)
                newsortDirection = SortDirection.Descending;
            else
                newsortDirection = SortDirection.Ascending;

            var newLst = newsortDirection == SortDirection.Ascending ? obj.AsQueryable().OrderBy(sortExpress).ToList() : obj.AsQueryable().OrderByDescending(sortExpress).ToList();
            HttpContext.Current.Session[sortExpression + obj.GetType()] = newsortDirection;
            return newLst;
        }
        #endregion
    }
}