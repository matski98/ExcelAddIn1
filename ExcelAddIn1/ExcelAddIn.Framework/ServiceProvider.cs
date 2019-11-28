using System.Collections.Generic;

namespace ExcelAddIn.Framework
{
    using System;

    public static class ServiceProvider
    {
        private static readonly object SyncRoot = new Object();
        private static volatile Dictionary<Type, object> _services;
        private static volatile Dictionary<Type, object> _servicesAction;

        private static Dictionary<Type, object> Services
        {
            get
            {
                if (_services == null)
                {
                    lock (SyncRoot)
                    {
                        if (_services == null)
                        {
                            _services = new Dictionary<Type, object>();
                        }
                    }
                }

                return _services;
            }
        }

        public static void RegisterService(Type type, object service, bool actionLifeCycle = false)
        {
            if (actionLifeCycle)
            {
                if (!ActionServices.ContainsKey(type))
                {
                    ActionServices.Add(type, service);
                }
            }
            else if (!Services.ContainsKey(type))
            {
                Services.Add(type, service);
            }
        }

        private static Dictionary<Type, object> ActionServices
        {
            get
            {
                if (_servicesAction == null)
                {
                    lock (SyncRoot)
                    {
                        if (_servicesAction == null)
                        {
                            _servicesAction = new Dictionary<Type, object>();
                        }
                    }
                }

                return _servicesAction;
            }
        }
        /*
        public static void RegisterService(Type type, object service)
        {
            if (!Services.ContainsKey(type))
            {
                Services.Add(type, service);
            }
        }*/

        public static T GetService<T>() where T : IService
        {
            return (T)GetService(typeof(T));
        }

        public static object GetService(Type serviceType)
        {
            return (Services.ContainsKey(serviceType)) ? Services[serviceType] : ((ActionServices.ContainsKey(serviceType)) ? ActionServices[serviceType] : null);
        }


        public static void ReleaseActionServices()
        {
            _servicesAction = null;
        }
    }
}
