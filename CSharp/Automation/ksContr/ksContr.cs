using System;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Steps.NET
{
	public class ksContr
	{
		public static void Main(string[] args)
		{
			System.Windows.Forms.Application.Run(ksContrForm.Instance);
		}
	}
}
