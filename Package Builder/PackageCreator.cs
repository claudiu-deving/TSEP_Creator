using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Tekla.Structures;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using System.Runtime.InteropServices;
using System.Reflection;


namespace TeklaTemplate.PackageBuilder
{
	/// <summary>
	/// Used at build to create the (.tsep) package file
	/// </summary>
	public class PackageCreator
	{
		/*
		//Needs
		//Folder definitions --- Done
		//Macro file to be used inside Tekla  --- Done
		//A way to figure out what assemblies are needed or if to just upload them all -- Just upload all of them, it is safer
		//Manifest generator to create the manifest.xml file -- Done
		//Start the packageBuilder batch process -- DONE
		//Build the package --DONE
		//Introduce the process creator locally -- DONE
		//Connect the current Tekla Model instace and create some of the variables upon it --DONE
		//Build the macro -- DONE
		//Connect it to the build part of Visual Studio --DONE -- It is connected to the debug part
		//Optimizations -- Fine for now

		*/
		private static readonly string teklaVersion = GetTeklaVersion(); //Tekla Version of the current instance
		private static readonly string componentGUID = "BE368DAD-8C59-4570-A764-6321C4F4DA11"; //Is static, one for each component
		private static readonly string executableName = System.AppDomain.CurrentDomain.FriendlyName; //e.g. "Test.exe'
		private static readonly bool isModelMacro = true; //This determines where to place the macro that calls the application
		private static readonly string toBeInstalledPackagesFolder = Path.Combine(GetEnvironmentFolder(), "Extensions", "To be installed"); //Where to temporary place the tsep file upon debugging
		private static readonly string installedPackagesFolder = Path.Combine(GetEnvironmentFolder(), "Extensions", "Installed"); //The directory where the extensions are installed
		private static readonly string macroFolder = Path.Combine(GetEnvironmentFolder(), "Environments", "common","macros","modeling"); //The macro folder
		private static readonly string currentDirectory = Environment.CurrentDirectory; //The directory of the bin folder
		private static readonly string ModelApplicationsTargetFolder = "ModelApplicationsTargetFolder"; 
		private static readonly string ModelPluginsTargetFolder = "ModelPluginsTargetFolder";
		private static readonly string ExtensionsTargetFolder = "ExtensionsTargetFolder";
		private static readonly string BinariesTargetFolder = "BinariesTargetFolder";
		private static readonly string BitmapsTargetFolder = "BitmapsTargetFolder";
		private static readonly string AttributeFileTargetFolder = "AttributeFileTargetFolder";
		private static readonly string CommonMacroTargetFolder = "CommonMacroTargetFolder";
		private static readonly string ModelMacroTargetFolder = "ModelMacroTargetFolder";
		private static readonly string DrawingMacroTargetFolder = "DrawingMacroTargetFolder";
		private static readonly string parentDirectory = Directory.GetParent(currentDirectory).Parent.FullName;
		private static readonly string resourcesFolder = Path.Combine(parentDirectory, "Resources");
		private static readonly string xmlFullPath = Path.Combine(parentDirectory, "Package Builder", "manifest.xml");
		private static readonly string binariesFolder = Directory.GetParent(Assembly.GetExecutingAssembly().Location).FullName;
		private static readonly string packageBuilderFolder = Path.Combine(parentDirectory,"Package Builder");
		private static readonly string packageBuilderMacroFolder = Path.Combine(parentDirectory, "Package Builder","Macro");
		private static readonly string assemblyName = typeof(Main.App).Assembly.GetName().Name;
		private static readonly GuidAttribute guid = (GuidAttribute)typeof(Main.App).Assembly.GetCustomAttributes(typeof(GuidAttribute), false).FirstOrDefault();
		private static readonly AssemblyFileVersionAttribute assemblyVersion = (AssemblyFileVersionAttribute)typeof(Main.App).Assembly.GetCustomAttributes(typeof(AssemblyFileVersionAttribute), false).FirstOrDefault();
		private static readonly AssemblyCompanyAttribute manufacturer = (AssemblyCompanyAttribute)typeof(Main.App).Assembly.GetCustomAttributes(typeof(AssemblyCompanyAttribute), false).FirstOrDefault();
		private static readonly AssemblyDescriptionAttribute assemblyDescription = (AssemblyDescriptionAttribute)typeof(Main.App).Assembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false).FirstOrDefault();
		private static readonly ManifestResourceInfo manifestResourceInfo = new ManifestResourceInfo(Assembly.GetExecutingAssembly(), "icon.ico", ResourceLocation.Embedded);


		/// <summary>
		/// Change the name of the Png file in the macro directory from Tekla Template.png to the current assembly name.png
		/// </summary>
		static void ChangePngName()
		{
			if (File.Exists(Path.Combine(packageBuilderMacroFolder, "Tekla Template.png")))
				System.IO.File.Move(Path.Combine(packageBuilderMacroFolder, "Tekla Template.png"), Path.Combine(packageBuilderMacroFolder, $"{assemblyName}.png"));
		}
		/// <summary>
		/// The reload macro for the applications catalog
		/// </summary>
		static void CreateAndRunReloadMacro()
		{
			string script = $@" #pragma warning disable 1633 // Unrecognized #pragma directive
								#pragma reference ""Tekla.Macros.Wpf.Runtime""
								#pragma reference ""Tekla.Macros.Runtime""
								#pragma warning restore 1633 // Unrecognized #pragma directive

								namespace UserMacros {{
									public sealed class Macro {{
										[Tekla.Macros.Runtime.MacroEntryPointAttribute()]
										public static void Run(Tekla.Macros.Runtime.IMacroRuntime runtime) {{
											Tekla.Macros.Wpf.Runtime.IWpfMacroHost wpf = runtime.Get<Tekla.Macros.Wpf.Runtime.IWpfMacroHost>();
											wpf.View(""CatalogTree.CatalogTreeView"").Find(""AdvancedMenu"", ""AdvancedMenu.Root"", ""AdvancedMenu.CatalogManagement"", ""AdvancedMenu.ReloadCatalogFilesAndTranslations"").As.Button.Invoke();
										}}
									}}
								}}
			";
			using (StreamWriter writer = new StreamWriter(Path.Combine(macroFolder , "Reload Catalog.cs")))
			{
				writer.Write(script);
				writer.Close();
			}
			Tekla.Structures.Model.Operations.Operation.RunMacro("Reload Catalog.cs");
		}

		/// <summary>
		/// Creates the macro file to be called from the Tekla UI
		/// </summary>
		static void CreateMacroFile()
		{
			string script = $@" using System.Diagnostics;
								using System.IO;
								using System.Windows.Forms;
								using Tekla.Structures;
								using TSM = Tekla.Structures.Model;
								using Tekla.Structures.Model;

								namespace Tekla.Technology.Akit.UserScript
								{{
									public class Script
									{{
										public static void Run(Tekla.Technology.Akit.IScript akit)
										{{
											string TSBinaryDir = """";
											TSM.Model CurrentModel = new TSM.Model();
											TeklaStructuresSettings.GetAdvancedOption(""XSDATADIR"", ref TSBinaryDir);
			
											string ApplicationName = ""{executableName}"";
											string ApplicationPath = Path.Combine(TSBinaryDir, ""Environments\\common\\extensions\\{assemblyName}\\"" + ApplicationName);

											Process NewProcess = new Process();

											if (File.Exists(ApplicationPath))
											{{
												NewProcess.StartInfo.FileName = ApplicationPath;

												try
												{{
													NewProcess.Start();
													NewProcess.WaitForExit();
												}}
												catch
												{{
													MessageBox.Show(ApplicationName + "" failed to start."");
												}}
											}}
											else
											{{
												MessageBox.Show(ApplicationName + "" not found."");
											}}
										}}
									}}
								}}
			";
			using (StreamWriter writer = new StreamWriter(packageBuilderMacroFolder+$@"\{assemblyName}.cs"))
			{
				writer.Write(script);
				writer.Close();
			}
		}

		/// <summary>
		/// Gets the environment folder
		/// </summary>
		/// <returns></returns>
		static string GetEnvironmentFolder()
		{
			string TSBinaryDir = "";
			TeklaStructuresSettings.GetAdvancedOption("XSDATADIR", ref TSBinaryDir);
			return TSBinaryDir;
		}


		/// <summary>
		/// Retrieves the current version of Tekla from a string of type "2020 Service Pack 10"
		/// </summary>
		/// <returns>A string of type 2020.0</returns>
		static string GetTeklaVersion()
		{
			string version = TeklaStructuresInfo.GetCurrentProgramVersion().Substring(0, 4) + ".0";
		   return version;
		}

		/// <summary>
		/// Rewrites the given string a safe way for the package builder to handle
		/// </summary>
		/// <param Name="text"></param>
		/// <returns></returns>
		private static string GetSafeID(string text)
		{
			text = Regex.Replace(text, @"[^\w\s]", " "); // allow only word characters (a-z, 0-9, _) and white space, change others to space
			text = Regex.Replace(text, @"\s+", "_"); // convert all adjacent spaces to one underscore
			return text;
		}

		/// <summary>
		/// Builds a Tsep type package using a specific folder format
		/// </summary>
		public static void BuildPackage()
		{
			ChangePngName();
			XMLNodes productNode = CreateProductNode();
			XMLNodes sourcePathVariables = CreateSourcePaths();
			XMLNodes targetPathVariables = CreateTargetPaths();
			XMLNodes component = CreateComponent();
			XMLNodes feature = CreateFeature(component);

			var listFirstChildren = new List<XMLNodes> { productNode, sourcePathVariables, targetPathVariables,component, feature };
			XMLNodes nodez = new XMLNodes("TEP", new List<Tuple<string, string>> { new Tuple<string, string>("Version", "2.0") }, listFirstChildren);
			ManifestXML manifestXML = new ManifestXML();
			manifestXML.CreateXML(nodez);
			CreateMacroFile();
			Build();
			IntroduceTheUninstallFile();
			Install();
			CopyPackageFile();
			Install();
			CreateAndRunReloadMacro();
		}

		/// <summary>
		/// Creates the components that hold the information about the files
		/// </summary>
		/// <summary>
		/// Source and Target path
		/// </summary>
		/// <returns>A XMLNodes type</returns>
		private static XMLNodes CreateComponent()
		{
			List<string> resourcesFolders = new List<string>() { resourcesFolder, binariesFolder, packageBuilderMacroFolder};
		   

			List<XMLNodes> ComponentNodes = new List<XMLNodes>();
			foreach (string folder in resourcesFolders)
			{
				if (folder == resourcesFolder)
				{
					foreach (string file in GetFiles(folder))
					{
						var fileName = file.Replace(folder + "\\", "");
						var node = XMLNodes.Component(fileName, $"%Resources%\\{fileName}", $"%BitmapsTargetFolder%");
						ComponentNodes.Add(node);
					}
				}
				if (folder == binariesFolder)
				{
					foreach (string file in GetFiles(folder, "*.dll"))
					{
						var fileName = file.Replace(folder + "\\", "");
						var node = XMLNodes.Component(fileName, $"%BinariesFolder%\\{fileName}", $"%BinariesTargetFolder%");
						ComponentNodes.Add(node);
					}
					foreach (string file in GetFiles(folder, $"{executableName}"))
					{
						var fileName = file.Replace(folder + "\\", "");
						var node = XMLNodes.Component(GetSafeID(fileName), $"%BinariesFolder%\\{fileName}", $"%BinariesTargetFolder%");
						ComponentNodes.Add(node);
					}
				}
				if (folder == packageBuilderMacroFolder)
				{
					foreach (string file in GetFiles(folder))
					{
						if (file.EndsWith(".cs"))
						{
							if (isModelMacro)
							{
								var fileName = file.Replace(folder + "\\", "");
								var node = XMLNodes.Component(GetSafeID(fileName), $"%MacroFolder%\\{fileName}", $"%ModelMacroTargetFolder%");
								ComponentNodes.Add(node);
							}
							else
							{
								var fileName = file.Replace(folder + "\\", "");
								var node = XMLNodes.Component(GetSafeID(fileName), $"%MacroFolder%\\{fileName}", $"%DrawingMacroTargetFolder%");
								ComponentNodes.Add(node);
							}
						}
						else if (file.EndsWith(".png"))
						{
							if (isModelMacro)
							{
								var fileName = file.Replace(folder + "\\", "");
								var node = XMLNodes.Component(GetSafeID(fileName), $"%MacroFolder%\\{fileName}", $"%ModelMacroTargetFolder%");
								ComponentNodes.Add(node);
							}
							else
							{
								var fileName = file.Replace(folder + "\\", "");
								var node = XMLNodes.Component(GetSafeID(fileName), $"%MacroFolder%\\{fileName}", $"%DrawingMacroTargetFolder%");
								ComponentNodes.Add(node);
							}
						};
					}
				}
			}

			XMLNodes component = new XMLNodes("Component",
			 new List<Tuple<string, string>> { new Tuple<string, string>("Id", "TheExtensionComponent"),
												  new Tuple<string, string>("Guid", componentGUID )}, ComponentNodes
			 );
			return component;
		}

		/// <summary>
		/// Creates the component Reference and introduces it in a Feature Node
		/// </summary>
		/// <param Name="component"></param>
		/// <returns></returns>
		private static XMLNodes CreateFeature(XMLNodes component)
		{
			XMLNodes componentRef = new XMLNodes("ComponentRef",
			 new List<Tuple<string, string>> { new Tuple<string, string>("ReferenceId", component.attributes[0].Item2) });


			XMLNodes feature = new XMLNodes("Feature",
			new List<Tuple<string, string>> { new Tuple<string, string>("Id", "TheExtensionFeature"),
											  new Tuple<string, string>("Title", "ExtensionFeature") }, new List<XMLNodes> { componentRef });
			return feature;
		}

		/// <summary>
		/// Gets the files in a folder at a given path with the given filter
		/// </summary>
		/// <param Name="path"></param>
		/// <param Name="filter">The filer (e.g. *.exe)</param>
		/// <returns></returns>
		static IEnumerable<string> GetFiles(string path, string filter =null)
		{
			Queue<string> queue = new Queue<string>();
			queue.Enqueue(path);
			while (queue.Count > 0)
			{
				path = queue.Dequeue();
				try
				{
					foreach (string subDir in Directory.GetDirectories(path))
					{
						queue.Enqueue(subDir);
					}
				}
				catch (Exception ex)
				{
					Console.Error.WriteLine(ex);
				}
				string[] files = null;
				try
				{
					if(filter!=null)
					files = Directory.GetFiles(path,filter);
					else files = Directory.GetFiles(path);
				}
				catch (Exception ex)
				{
					Console.Error.WriteLine(ex);
				}
				if (files != null)
				{
					for (int i = 0; i < files.Length; i++)
					{
						yield return files[i];
					}
				}
			}
		}

		/// <summary>
		/// Creates the parent node for all target paths nodes
		/// </summary>
		/// <returns>A XMLNodes with all the Children</returns>
		private static XMLNodes CreateTargetPaths()
		{
			//Modify here locally any changes 
			var ModelPluginsDirectory = XMLNodes.TargetPathVariable(ModelPluginsTargetFolder, "%ENVDIR%\\extensions\\plugins\\tekla\\model\\");
			var ModelApplicationsDirectory = XMLNodes.TargetPathVariable(ModelApplicationsTargetFolder, "%ENVDIR%\\extensions\\applications\\tekla\\model\\");
			var ExtensionsDir = XMLNodes.TargetPathVariable(ExtensionsTargetFolder, "%commonEnvFolder%\\extensions\\");
			var BinariesTargetDirectory = XMLNodes.TargetPathVariable(BinariesTargetFolder, $"%ExtensionsTargetFolder%\\{assemblyName}\\");
			var BitmapsDirectory = XMLNodes.TargetPathVariable(BitmapsTargetFolder, "%ENVDIR%\\..\\bitmaps\\");
			var AttributeFileDirectory = XMLNodes.TargetPathVariable(AttributeFileTargetFolder, "%commonEnvFolder%\\system\\");
			var CommonMacroDirectory = XMLNodes.TargetPathVariable(CommonMacroTargetFolder, "%commonEnvFolder%\\macros\\");
			var ModelMacroDirectory = XMLNodes.TargetPathVariable(ModelMacroTargetFolder, "%commonEnvFolder%\\macros\\modeling\\");
			var DrawinglMacroDirectory = XMLNodes.TargetPathVariable(DrawingMacroTargetFolder, "%commonEnvFolder%\\macros\\drawings\\");
			XMLNodes TargetPathVariables = new XMLNodes("TargetPathVariables", null, new List<XMLNodes> { ModelPluginsDirectory,
																										  ModelApplicationsDirectory,
																										  ExtensionsDir,
																										  BinariesTargetDirectory,
																										  BitmapsDirectory,
																										  AttributeFileDirectory,
																										  CommonMacroDirectory,
																										  ModelMacroDirectory,
																										  DrawinglMacroDirectory
			});
			return TargetPathVariables;
		}

		/// <summary>
		/// Creates the parent node for all source paths nodes
		/// </summary>
		/// <returns>A XMLNodes with all the Children</returns>
		private static XMLNodes CreateSourcePaths()
		{
			var OutPutFolder = XMLNodes.SourcePathVariable("TepOutputFolder", "%TEPDEFINITIONFILEFOLDER%\\output" + "\\");
			var BinariesFolder = XMLNodes.SourcePathVariable("BinariesFolder", binariesFolder);
			var BitmapsFolder = XMLNodes.SourcePathVariable("Resources", resourcesFolder);
			var MacroFolder = XMLNodes.SourcePathVariable("MacroFolder", packageBuilderMacroFolder);
			XMLNodes SourcePathVariables = new XMLNodes("SourcePathVariables", null, new List<XMLNodes> { OutPutFolder, BinariesFolder, BitmapsFolder,MacroFolder });
			return SourcePathVariables;
		}

		/// <summary>
		/// Creates the product node with all its Children
		/// </summary>
		/// <returns></returns>
		private static XMLNodes CreateProductNode()
		{
			var Id = new Tuple<string, string>("Id", GetSafeID(assemblyName));
			var UpgradeCode = new Tuple<string, string>("UpgradeCode", guid.Value.ToString());
			var Version = new Tuple<string, string>("Version", assemblyVersion.Version);
			var Language = new Tuple<string, string>("Language", "1033");
			var Name = new Tuple<string, string>("Name", assemblyName);
			var Manufacturer = new Tuple<string, string>("Manufacturer", manufacturer.Company);
			var Description = new Tuple<string, string>("Description", assemblyDescription.Description);
			var IconPath = new Tuple<string, string>("IconPath", Path.Combine(resourcesFolder, manifestResourceInfo.FileName));
			XMLNodes TeklaVersion = new XMLNodes("TeklaVersion", new List<Tuple<string, string>> { new Tuple<string, string>("Name", "2099.1") });

			XMLNodes MinTeklaVersion = new XMLNodes("MinTeklaVersion", new List<Tuple<string, string>> { new Tuple<string, string>("Name", "2016.1") });

			XMLNodes MaxTeklaVersion = new XMLNodes("MaxTeklaVersion", new List<Tuple<string, string>> { new Tuple<string, string>("Name", "2099.1") });

			XMLNodes TeklaVersions = new XMLNodes("TeklaVersions", null, new List<XMLNodes> { TeklaVersion, MinTeklaVersion, MaxTeklaVersion });

			XMLNodes productNode = new XMLNodes("Product",
											   new List<Tuple<string, string>> {
													Id,
													UpgradeCode,
													Version,
													Language,
													Name,
													Manufacturer,
													Description,
													IconPath  }, new List<XMLNodes> { TeklaVersions });
			return productNode;
		}

		/// <summary>
		/// Calls the BatchBuilder executable and builds the package with following the XML File rules
		/// </summary>
		static void Build()
		{
			ProcessStartInfo ProcessInfo = new ProcessStartInfo();
			ProcessInfo.Arguments = $"-i \"{xmlFullPath}\" -o \"{packageBuilderFolder}\\{assemblyName+"."+assemblyVersion.Version}.tsep\" -l \"{packageBuilderFolder}\"";
			ProcessInfo.FileName = $@"C:\Program Files\Tekla Structures\{teklaVersion}\nt\bin\TeklaExtensionPackage.BatchBuilder.exe";
			Process process;
			ProcessInfo.CreateNoWindow = false;
			ProcessInfo.UseShellExecute = false;
			ProcessInfo.RedirectStandardError = true;
			ProcessInfo.RedirectStandardOutput = true;
			process = Process.Start(ProcessInfo);
			process.WaitForExit();
			process.Close();
		}

		/// <summary>
		/// <para>
		/// Uninstalls the tsep files with the RemoveExtensionOnStartup file in them
		/// </para>
		/// <para>
		/// Installs the extensions from the ToBeInstalled folder
		/// </para>
		/// </summary>
		public static void Install()
		{
			ProcessStartInfo ProcessInfo = new ProcessStartInfo();
			ProcessInfo.Arguments = $"\"2020.0\" \"C:\\ProgramData\\Trimble\\Tekla Structures\\2020.0\" \"2020\"";
			ProcessInfo.FileName = $@"C:\Program Files\Tekla Structures\2020.0\nt\bin\TeklaExtensionPackage.TepAutoInstaller.exe";
			Process process;
			ProcessInfo.CreateNoWindow = true;
			ProcessInfo.UseShellExecute = false;
			ProcessInfo.RedirectStandardError = true;
			ProcessInfo.RedirectStandardOutput = true;
			process = Process.Start(ProcessInfo);
			process.WaitForExit();
			process.Close();
		}

		/// <summary>
		/// Copies the locally created tsep file to the ToBeInstalled folder
		/// </summary>
		static void CopyPackageFile()
		{

			string sourceFile = $@"{packageBuilderFolder}\{assemblyName + "." + assemblyVersion.Version}.tsep";
			string destFile = toBeInstalledPackagesFolder+"\\"+$"{ assemblyName + "." + assemblyVersion.Version}.tsep";
			System.IO.File.Copy(sourceFile, destFile, true);
		}

		/// <summary>
		/// Introduces the tsep file to be unistalled upon called the AutoInstaller
		/// </summary>
	   static void IntroduceTheUninstallFile()
		{
			string destFile = installedPackagesFolder + "\\" + "{"+$"{assemblyName}"+"}" +"{" + $"{assemblyVersion.Version}" +"}"+"{"+$"{guid.Value}"+"}";
			Console.WriteLine(Path.Combine(destFile, "RemoveExtensionOnStartup"));
			if(Directory.Exists(Path.Combine(destFile)))
			{
				File.WriteAllText(Path.Combine(destFile, "RemoveExtensionOnStartup"), "");
			}
		}


	}

	/// <summary>
	/// The class that builds up the manifest xml file used for the package creation
	/// </summary>
	internal class ManifestXML 
	{
	   public  void CreateXML(XMLNodes nodes)
	   {
			XmlDocument XmlDoc = new XmlDocument();
			string xmlFullPath = Path.Combine(Directory.GetParent(Environment.CurrentDirectory).Parent.FullName, "Package Builder", "manifest.xml");  
			XmlNode rootNode = XmlDoc.CreateElement(nodes.Name);
			XmlDoc.AppendChild(rootNode);
			XmlAttribute versionAttr = XmlDoc.CreateAttribute(nodes.attributes[0].Item1);
			versionAttr.Value = nodes.attributes[0].Item2;
			rootNode.Attributes.Append(versionAttr);
			XmlNode xmlNode ;
			foreach (var elements in nodes.Children)
			{
			  xmlNode = XmlDoc.CreateElement(elements.Name);
			  rootNode.AppendChild(xmlNode);
				if (elements.attributes !=null)
				{
					foreach (var attribute in elements.attributes)
					{
						var xmlAttribute = XmlDoc.CreateAttribute(attribute.Item1);
						xmlAttribute.Value = attribute.Item2;
					   xmlNode.Attributes.Append(xmlAttribute);
					}
				}
				if (elements.Children !=null)
				{
					foreach (var child in elements.Children)
					{
						XmlNode xmlChild = XmlDoc.CreateElement(child.Name);
						xmlNode.AppendChild(xmlChild);
						if (child.attributes != null)
						{
							foreach (var attribute in child.attributes)
							{
								var xmlAttribute = XmlDoc.CreateAttribute(attribute.Item1);
								xmlAttribute.Value = attribute.Item2;
								xmlChild.Attributes.Append(xmlAttribute);
							}
						}

						if (child.Children != null)
						{

							foreach (var grandChild in child.Children)
							{
								  XmlNode xmlGrandChild = XmlDoc.CreateElement(grandChild.Name);
								  xmlChild.AppendChild(xmlGrandChild);
								if (grandChild.attributes != null)
								{
									foreach (var attribute in grandChild.attributes)
									{
										  var xmlAttribute = XmlDoc.CreateAttribute(attribute.Item1);
										  xmlAttribute.Value = attribute.Item2;
										 xmlGrandChild.Attributes.Append(xmlAttribute);
									}
								}
							}
						}
					}
				}
			}
			
			XmlTextWriter xmlWriter = new XmlTextWriter(xmlFullPath, Encoding.UTF8)
			{
				Formatting = Formatting.Indented
			};
			xmlWriter.WriteStartDocument();
			XmlDoc.WriteContentTo(xmlWriter);
			xmlWriter.WriteEndDocument();
			xmlWriter.Close();
		}
	}

	/// <summary>
	/// Represent a XML Node with Name, its attributes and it's Children
	/// </summary>
	internal class XMLNodes 
	{
		/// <summary>
		/// The Name of the node e.g. <summary></summary>
		/// </summary>
		public  string Name { get; set; }

		/// <summary>
		/// The list of attributes with Name and value (e.g. Id="Some_unique_id",Value="SomeValue")
		/// </summary>
		public  List<Tuple<string, string>> attributes { get; set; }

		/// <summary>
		/// The constructor for the node, it can miss attributes or Children
		/// </summary>
		/// <param Name="Name"></param>
		/// <param Name="Attributes"></param>
		/// <param Name="Children"></param>
		public  XMLNodes(string Name, List<Tuple<string, string>> Attributes = null, List<XMLNodes> Children=null)
		{
			this.Name = Name;
			attributes = Attributes;
			this.Children = Children;
		}

		/// <summary>
		/// List of Children of the node
		/// </summary>
		public  List<XMLNodes> Children {get;set; }

		/// <summary>
		/// Specific type of Xml node 
		/// </summary>
		/// <param Name="Name"></param>
		/// <param Name="Value"></param>
		/// <returns></returns>
		public static XMLNodes SourcePathVariable(string Name,string Value)
		{
			return new XMLNodes("SourcePathVariable", new List<Tuple<string, string>>() { new Tuple<string, string>("Id", Name), new Tuple<string, string>("Value", Value) });
		}

		/// <summary>
		/// Specific type of Xml node 
		/// </summary>
		/// <param Name="Name"></param>
		/// <param Name="Value"></param>
		/// <returns></returns>
		public static XMLNodes TargetPathVariable(string Name, string Value)
		{
			return new XMLNodes("PathVariable", new List<Tuple<string, string>>() { new Tuple<string, string>("Id", Name), new Tuple<string, string>("Value", Value) });
		}

		/// <summary>
		/// Specific type of Xml node 
		/// </summary>
		/// <param Name="Name"></param>
		/// <param Name="Source"></param>
		/// <param Name="Target"></param>
		/// <returns></returns>
		public static XMLNodes Component(string Name, string Source,string Target) 
		{
			return new XMLNodes("File", new List<Tuple<string, string>>() {   new Tuple<string, string>("Id", Name), 
																							new Tuple<string, string>("Source", Source),
																							new Tuple<string, string>("Target", Target),
			});

		}

	 
	}
}
