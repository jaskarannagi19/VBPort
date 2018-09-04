using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;

namespace CSharpInteropService
{
    [ComVisible(true), Guid(LibraryInvoke.EventsId), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ILibraryInvokeEvent
    {
        [DispId(1)]
        void MessageEvent(string message);
    }

    [ComVisible(true), Guid(LibraryInvoke.InterfaceId)]
    public interface ILibraryInvoke
    {
        [DispId(1)]
        object[] GenericInvoke(string dllFile, string className, string methodName, object[] parameters);
    }

    [ComVisible(true), Guid(LibraryInvoke.ClassId)]
    [ComSourceInterfaces("CSharpInteropService.ILibraryInvokeEvent")]
    [ComClass(LibraryInvoke.ClassId, LibraryInvoke.InterfaceId, LibraryInvoke.EventsId)]
    public class LibraryInvoke : ILibraryInvoke
    {
        public const string ClassId = "3D853E7B-01DA-4944-8E65-5E36B501E889";
        public const string InterfaceId = "CB344AD3-88B2-47D8-86F1-20EEFAF6BAE8";
        public const string EventsId = "5E16F11C-2E1D-4B35-B190-E752B283260A";

        public delegate void MessageHandler(string message);
        public event MessageHandler MessageEvent;

        public object[] GenericInvoke(string dllFile, string className, string methodName, object[] parameters)
        {
            Assembly dll = Assembly.LoadFrom(dllFile);

            Type classType = dll.GetType(className);
            object classInstance = Activator.CreateInstance(classType);
            MethodInfo classMethod = classType.GetMethod(methodName);

            EventInfo eventMessageEvent = classType.GetEvent("MessageEvent", BindingFlags.NonPublic | BindingFlags.Static);

            if (eventMessageEvent != null)
            {
                Type typeMessageEvent = eventMessageEvent.EventHandlerType;

                MethodInfo handler = typeof(LibraryInvoke).GetMethod("OnMessageEvent", BindingFlags.NonPublic | BindingFlags.Instance);
                Delegate del = Delegate.CreateDelegate(typeMessageEvent, this, handler);

                MethodInfo addHandler = eventMessageEvent.GetAddMethod(true);
                Object[] addHandlerArgs = { del };
                addHandler.Invoke(classInstance, addHandlerArgs);
            }

            return (object[])classMethod.Invoke(classInstance, parameters);
        }

        private void OnMessageEvent(string message)
        {
            MessageEvent?.Invoke(message);
        }
    }
}
