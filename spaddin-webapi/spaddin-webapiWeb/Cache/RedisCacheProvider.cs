using Newtonsoft.Json;
using StackExchange.Redis;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace spaddin_webapiWeb.Cache
{
    public class RedisCacheProvider : OfficeDevPnP.Core.Utilities.Cache.ICacheProvider
    {
        #region Constants
        public const string RedisConnectionStringKey = "REDIS";
        #endregion

        #region Fields
        private Lazy<ConnectionMultiplexer> _lazyConnection;
        #endregion

        #region Ctors
        public RedisCacheProvider(string connectionString=null)
        {
            connectionString = connectionString ?? ConfigurationManager.ConnectionStrings[RedisConnectionStringKey].ConnectionString;
            _lazyConnection = new Lazy<ConnectionMultiplexer>(() =>
            {
                return ConnectionMultiplexer.Connect(connectionString);
            });
        }
        #endregion

        #region Private Properties
        private ConnectionMultiplexer Connection
        {
            get
            {
                return _lazyConnection.Value;
            }
        }
        #endregion

        #region Cache API
        public T Get<T>(string cacheKey)
        {
            try
            {
                IDatabase cache = Connection.GetDatabase();
                return JsonConvert.DeserializeObject<T>(cache.StringGet(cacheKey));
            }
            catch (Exception)
            {
                throw new Exception($"An unexpected error occured while trying to fetch value from cache with key ${cacheKey}");
            }
        }

        public void Put<T>(string cacheKey, T item)
        {
            try
            {
                IDatabase cache = Connection.GetDatabase();
                cache.StringSet(cacheKey, JsonConvert.SerializeObject(item));
            }
            catch (Exception)
            {
                throw new Exception($"An unexpected error occured while trying to cache value for key ${cacheKey}");
            }
        }
        #endregion
    }
}