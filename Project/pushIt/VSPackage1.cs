using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

namespace pushIt
{
	/// <summary>
	/// This is the class that implements the package exposed by this assembly.
	/// </summary>
	/// <remarks>
	/// <para>
	/// The minimum requirement for a class to be considered a valid package for Visual Studio
	/// is to implement the IVsPackage interface and register itself with the shell.
	/// This package uses the helper classes defined inside the Managed Package Framework (MPF)
	/// to do it: it derives from the Package class that provides the implementation of the
	/// IVsPackage interface and uses the registration attributes defined in the framework to
	/// register itself and its components with the shell. These attributes tell the pkgdef creation
	/// utility what data to put into .pkgdef file.
	/// </para>
	/// <para>
	/// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
	/// </para>
	/// </remarks>
	[PackageRegistration(UseManagedResourcesOnly = true)]
	[InstalledProductRegistration("#110",
		"#112",
		"1.0",
		IconResourceID = 400)] // Info on this package for Help/About
	[Guid(VSPackage1.PackageGuidString)]
	[ProvideAutoLoad(UIContextGuids80
		.SolutionExists)]
	[SuppressMessage("StyleCop.CSharp.DocumentationRules",
		"SA1650:ElementDocumentationMustBeSpelledCorrectly",
		Justification = "pkgdef," + "VS and vsixmanifest are valid VS terms")]
	public sealed class VSPackage1 : Package
	{
		/// <summary>
		/// VSPackage1 GUID string.
		/// </summary>
		public const string PackageGuidString = "d41d0918-8b0f-49b3-b43d-a882b68e7ae6";

		/// <summary>
		/// Initializes a new instance of the <see cref="VSPackage1"/> class.
		/// </summary>
		public VSPackage1()
		{
			// Inside this method you can place any initialization code that does not require
			// any Visual Studio service because at this point the package object is created but
			// not sited yet inside Visual Studio environment. The place to do all the other
			// initialization is the Initialize method.
		}

		#region Package Members

		/// <summary>
		/// Initialization of the package; this method is called right after the package is sited, so this is the place
		/// where you can put all the initialization code that rely on services provided by VisualStudio.
		/// </summary>
		protected override void Initialize()
		{
			base
				.Initialize();

			DTE dte = Package
				.GetGlobalService(typeof(DTE)) as DTE;

			dte
				.Events
				.SolutionEvents
				.BeforeClosing += SolutionEvents_BeforeClosing;
		}

		private void SolutionEvents_BeforeClosing()
		{
			if (!(GetService(typeof(IVsSolution)) is IVsSolution solution))
				return;

			if (!(solution
				.GetSolutionInfo(out var solutionDirectory,
				out _,
				out _) == VSConstants
				.S_OK && !String
				.IsNullOrEmpty(solutionDirectory)))
				return;

			var pushBat = Path
				.Combine(solutionDirectory,
				"!push.bat");

			if (!File
				.Exists(pushBat))
				return;

			MessageBox
				.Show(pushBat);

			return;

			ProcessStartInfo startInfo1 = new ProcessStartInfo(pushBat);

			startInfo1
				.UseShellExecute = true;

			System
				.Diagnostics
				.Process
				.Start(startInfo1);
		}

		#endregion
	}
}
