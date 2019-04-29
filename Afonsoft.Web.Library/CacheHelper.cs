using System;
using System.Collections;
using System.Web;

namespace Afonsoft.Web.Library
{
    /// <summary>
    /// Classe por manipular o Cache compatilhado entre aplicações.
    /// </summary>
    public static class CacheHelper
    {
        /// <summary>
        /// Insert value into the cache using
        /// appropriate name/value pairs
        /// </summary>
        /// <typeparam name="T">Type of cached item</typeparam>
        /// <param name="o">Item to be cached</param>
        /// <param name="key">Name of item</param>
        public static void Add<T>(string key, T o)
        {
            if (Exists(key))
                Clear(key);

            // In this example, I want an absolute
            // timeout so changes will always be reflected
            // at that time. Hence, the NoSlidingExpiration.
            if (HttpContext.Current != null)
                HttpContext.Current.Cache.Insert(key, o, null, DateTime.Now.AddMinutes(30), System.Web.Caching.Cache.NoSlidingExpiration);

            if (HttpContext.Current != null && HttpContext.Current.Session != null)
                HttpContext.Current.Session[key] = o;
        }
        /// <summary>
        /// Insert value into the cache using
        /// appropriate name/value pairs
        /// </summary>
        /// <typeparam name="T">Type of cached item</typeparam>
        /// <param name="o">Item to be cached</param>
        /// <param name="key">Name of item</param>
        /// <param name="timeExpiration">Time to expiration</param>
        public static void Add<T>(string key, T o, DateTime timeExpiration)
        {
            if (Exists(key))
                Clear(key);

            // In this example, I want an absolute
            // timeout so changes will always be reflected
            // at that time. Hence, the NoSlidingExpiration.

            if (HttpContext.Current != null)
                HttpContext.Current.Cache.Insert(key, o, null, timeExpiration.AddMinutes(1), System.Web.Caching.Cache.NoSlidingExpiration);

            if (HttpContext.Current != null && HttpContext.Current.Session != null)
                HttpContext.Current.Session[key] = o;

        }

        /// <summary>
        /// Remove item from cache
        /// </summary>
        /// <param name="key">Name of cached item</param>
        public static void Clear(string key)
        {
            if (HttpContext.Current != null)
            {
                HttpContext.Current.Cache.Remove(key);
                if (HttpContext.Current.Session != null)
                    HttpContext.Current.Session[key] = null;
            }
        }


        /// <summary>
        /// Remove all item from cache
        /// </summary>
        public static void ClearAll()
        {
            if (HttpContext.Current != null)
            {
                IDictionaryEnumerator cacheEnumerator = HttpContext.Current.Cache.GetEnumerator();
                while (cacheEnumerator.MoveNext())
                    HttpContext.Current.Cache.Remove(cacheEnumerator.Key.ToString());
            }
        }

        /// <summary>
        /// Check for item in cache
        /// </summary>
        /// <param name="key">Name of cached item</param>
        /// <returns></returns>
        public static bool Exists(string key)
        {
            bool exist = false;
            if (HttpContext.Current != null)
            {
                exist = HttpContext.Current.Cache[key] != null;
                if (exist == false)
                    if (HttpContext.Current.Session != null)
                        exist = HttpContext.Current.Session[key] != null;
            }
            return exist;
        }

        /// <summary>
        /// Retrieve cached item
        /// </summary>
        /// <typeparam name="T">Type of cached item</typeparam>
        /// <param name="key">Name of cached item</param>
        /// <param name="value">Cached value. Default(T) if
        /// item doesn't exist.</param>
        /// <returns>Cached item as type</returns>
        public static bool Get<T>(string key, out T value)
        {
            try
            {
                if (!Exists(key))
                {
                    value = default(T);
                    return false;
                }

                if (HttpContext.Current.Session != null && HttpContext.Current.Session[key] != null)
                    value = (T)HttpContext.Current.Session[key];
                else if (HttpContext.Current != null)
                    value = (T)HttpContext.Current.Cache[key];
                else
                {
                    value = default(T);
                    return false;
                }
            }
            catch
            {
                value = default(T);
                return false;
            }

            return true;
        }
    }
}
