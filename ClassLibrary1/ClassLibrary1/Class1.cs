using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace ClassLibrary1
{
    public class Class1
    {
        [DllImport("user32.dll")]
        public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        public object[] ShowForm(bool param1, int param2, double param3, string param4)
        {
            Form1 form = new Form1(param1, param2, param3, param4);
            form.ShowDialog();
            return new object[] { 1, 2.5, "A" };
        }

        public void ShowFormParent(long parent)
        {
            Form1 form = new Form1();
            form.Show();

            IntPtr p = new IntPtr(parent);
            SetParent(form.Handle, p);
            
        }

        private delegate void MessageHandler(string message);
        private static event MessageHandler MessageEvent = delegate { };

        public static void OnMessageEvent(string message)
        {
            MessageEvent(message);
        }

    }
}
