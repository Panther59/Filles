// <copyright file="ProjectUtilities.cs" company="Ayvan">
// Copyright (c) 2020 All Rights Reserved
// </copyright>
// <author>UTKARSHLAPTOP\Utkarsh</author>
// <date>2020-03-28</date>

namespace VSUnitTestMaster.Utilities
{
	using System;
	using System.IO;
	using System.IO.Compression;
	using System.Threading.Tasks;

	/// <summary>
	/// Defines the <see cref="ProjectUtilities" />.
	/// </summary>
	public class ProjectUtilities
	{
		/// <summary>
		/// Defines the codeCoveragePath.
		/// </summary>
		private string codeCoveragePath;

		/// <summary>
		/// Defines the reportGeneratorPath.
		/// </summary>
		private string reportGeneratorPath;

		/// <summary>
		/// Defines the resourcePresent.
		/// </summary>
		private bool resourcePresent = false;

		/// <summary>
		/// The ExtractResourceAsync.
		/// </summary>
		/// <returns>The <see cref="Task"/>.</returns>
		public async Task ExtractResourceAsync()
		{
			await Task.Run(() =>
			{
				try
				{
					if (this.resourcePresent == false)
					{
						var rootPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
						var dllFileInfo = new FileInfo(rootPath);
						var resourceFolder = Path.Combine(dllFileInfo.DirectoryName, "Resources");
						this.codeCoveragePath = Path.Combine(resourceFolder, "CodeCoverage");
						if (Directory.Exists(Path.Combine(resourceFolder, "CodeCoverage")) == false)
						{
							ZipFile.ExtractToDirectory(Path.Combine(resourceFolder, "CodeCoverage.zip"), this.codeCoveragePath);
						}

						this.reportGeneratorPath = Path.Combine(resourceFolder, "ReportGenerator");
						if (Directory.Exists(Path.Combine(resourceFolder, "ReportGenerator")) == false)
						{
							ZipFile.ExtractToDirectory(Path.Combine(resourceFolder, "ReportGenerator.zip"), this.reportGeneratorPath);
						}

						this.resourcePresent = true;
					}
				}
				catch (Exception ex)
				{
				}
			});
		}

		/// <summary>
		/// The GenereteCodeCoverage.
		/// </summary>
		/// <param name="solutionPath">The solutionPath<see cref="string"/>.</param>
		/// <returns>The <see cref="string"/>.</returns>
		public async Task<string> GenerateCodeCoverageForSolutionAsync(string solutionPath)
		{
			return await Task.Run(async () =>
			{
				if (!resourcePresent)
				{
					await this.ExtractResourceAsync();
				}

				if (!this.resourcePresent)
				{
					throw new Exception("Unable to find Required component oon your machine");
				}

				var tempDirectoy = Path.Combine(Path.GetTempPath(), "Code Coverage");
				var projectReportBaseDirctory = Path.Combine(solutionPath, "Code Coverage");

				var directory = new DirectoryInfo(projectReportBaseDirctory);
				if (directory.Exists)
				{
					directory.Delete(true);
				}

				var testCommand = $"dotnet test \"{solutionPath}\" --results-directory:\"{tempDirectoy}\" --collect:\"Code Coverage\"";
				this.RunCMDCommand(testCommand);
				var files = Directory.GetFiles(tempDirectoy, "*.coverage", SearchOption.AllDirectories);
				if (files.Length == 0)
				{
					throw new Exception("Failed to create test covarage report. Make sure there are unit test present under selected directory");
				}

				var coveragePath = Path.Combine(projectReportBaseDirctory, "Report", "Report.coverage");
				foreach (var file in files)
				{
					directory = new DirectoryInfo(Path.Combine(projectReportBaseDirctory, "Report"));
					if (!directory.Exists)
					{
						directory.Create();
					}

					if (File.Exists(coveragePath))
					{
						File.Delete(coveragePath);
					}

					File.Move(file, coveragePath);
					break;
				}

				Directory.Delete(tempDirectoy, true);

				var coverageXmlPath = Path.Combine(projectReportBaseDirctory, "Report", "Report.coveragexml");

				if (File.Exists(coverageXmlPath))
				{
					File.Delete(coverageXmlPath);
				}
				var codeCoverageExepath = Path.Combine(this.codeCoveragePath, "CodeCoverage.exe");
				var analyzeCommand = $"analyze /output:\"{coverageXmlPath}\" \"{coveragePath}\"";
				this.RunCommand(codeCoverageExepath, analyzeCommand);

				if (files.Length == 0)
				{
					throw new Exception("Failed to parse coverage report.");
				}

				var reportgeneratorPath = Path.Combine(this.reportGeneratorPath, "ReportGenerator.dll");
				var reportGenerateCommand = $"dotnet \"{reportgeneratorPath}\" \"-reports:{coverageXmlPath}\" \"-targetdir:{projectReportBaseDirctory}\"";
				this.RunCMDCommand(reportGenerateCommand);

				var htmlPath = Path.Combine(projectReportBaseDirctory, "index.htm");
				if (File.Exists(htmlPath))
				{
					return htmlPath;
				}
				else
				{
					throw new Exception("Failed to generate view for code coverage report");
				}
			});
		}

		/// <summary>
		/// The RunCMDCommand.
		/// </summary>
		/// <param name="command">The command<see cref="string"/>.</param>
		private void RunCMDCommand(string command)
		{
			var process = new System.Diagnostics.Process();
			process.StartInfo.FileName = "C:\\Windows\\System32\\cmd.exe";
			process.StartInfo.Arguments = "/c " + command;
			process.StartInfo.UseShellExecute = false;
			process.StartInfo.CreateNoWindow = true;
			process.StartInfo.RedirectStandardOutput = true;
			process.OutputDataReceived += (sender, data) =>
			{
				Console.WriteLine(data.Data);
			};
			process.StartInfo.RedirectStandardError = true;
			process.ErrorDataReceived += (sender, data) =>
			{
				Console.WriteLine(data.Data);
			};

			process.Start();
			process.WaitForExit();
			process.Dispose();
		}

		/// <summary>
		/// The RunCommand.
		/// </summary>
		/// <param name="fileName">The fileName<see cref="string"/>.</param>
		/// <param name="command">The command<see cref="string"/>.</param>
		private void RunCommand(string fileName, string command)
		{
			var process = new System.Diagnostics.Process();
			process.StartInfo.FileName = fileName;
			process.StartInfo.Arguments = command;
			process.StartInfo.UseShellExecute = false;
			process.StartInfo.CreateNoWindow = true;
			process.StartInfo.RedirectStandardOutput = true;
			process.OutputDataReceived += (sender, data) =>
			{
				Console.WriteLine(data.Data);
			};
			process.StartInfo.RedirectStandardError = true;
			process.ErrorDataReceived += (sender, data) =>
			{
				Console.WriteLine(data.Data);
			};

			process.Start();
			process.WaitForExit();
			process.Dispose();
		}
	}
}
