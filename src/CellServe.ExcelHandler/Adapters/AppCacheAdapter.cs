using LazyCache;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;
using System.Text;
using System.Threading.Tasks;

namespace CellServe.ExcelHandler.Adapters
{
    /// <summary>
    /// Wraps up CachingService implementation to give it only one public constructor
    /// </summary>
    public class AppCacheAdapter : IAppCache
    {
        private readonly CachingService _cachingService;
        public ObjectCache ObjectCache => _cachingService.ObjectCache;

        public AppCacheAdapter()
        {
            _cachingService = new CachingService();
        }

        public void Add<T>(string key, T item)
        {
            _cachingService.Add<T>(key, item);
        }

        public void Add<T>(string key, T item, DateTimeOffset absoluteExpiration)
        {
            _cachingService.Add<T>(key, item, absoluteExpiration);
        }

        public void Add<T>(string key, T item, TimeSpan slidingExpiration)
        {
            _cachingService.Add<T>(key, item, slidingExpiration);
        }

        public void Add<T>(string key, T item, CacheItemPolicy policy)
        {
            _cachingService.Add<T>(key, item, policy);
        }

        public T Get<T>(string key)
        {
            return _cachingService.Get<T>(key);
        }

        public Task<T> GetAsync<T>(string key)
        {
            return _cachingService.GetAsync<T>(key);
        }

        public T GetOrAdd<T>(string key, Func<T> addItemFactory)
        {
            return _cachingService.GetOrAdd<T>(key, addItemFactory);
        }

        public T GetOrAdd<T>(string key, Func<T> addItemFactory, DateTimeOffset absoluteExpiration)
        {
            return _cachingService.GetOrAdd<T>(key, addItemFactory, absoluteExpiration);
        }

        public T GetOrAdd<T>(string key, Func<T> addItemFactory, TimeSpan slidingExpiration)
        {
            return _cachingService.GetOrAdd<T>(key, addItemFactory, slidingExpiration);
        }

        public T GetOrAdd<T>(string key, Func<T> addItemFactory, CacheItemPolicy policy)
        {
            return _cachingService.GetOrAdd<T>(key, addItemFactory, policy);
        }

        public Task<T> GetOrAddAsync<T>(string key, Func<Task<T>> addItemFactory)
        {
            return _cachingService.GetOrAddAsync<T>(key, addItemFactory);
        }

        public Task<T> GetOrAddAsync<T>(string key, Func<Task<T>> addItemFactory, CacheItemPolicy policy)
        {
            return _cachingService.GetOrAddAsync<T>(key, addItemFactory, policy);
        }

        public Task<T> GetOrAddAsync<T>(string key, Func<Task<T>> addItemFactory, DateTimeOffset expires)
        {
            return _cachingService.GetOrAddAsync<T>(key, addItemFactory, expires);
        }

        public Task<T> GetOrAddAsync<T>(string key, Func<Task<T>> addItemFactory, TimeSpan slidingExpiration)
        {
            return _cachingService.GetOrAddAsync<T>(key, addItemFactory, slidingExpiration);
        }

        public void Remove(string key)
        {
            _cachingService.Remove(key);
        }
    }
}
