#region Using Directives
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
#endregion

namespace StrataWrapper
{
    public sealed class Manager : IDisposable
    {
        #region Members
        private Version m_VersionIKVM;
        private Version m_VersionStrata;
        private readonly String m_DirectoryDeployment;
        private readonly String m_DirectoryIKVM;
        private readonly String m_DirectoryStrata;
        #endregion

        #region Members (Static)
        private static readonly Regex s_RegexCompilerVersion = new Regex(@"^StrataWrapper\.IKVM\-(\d+.\d+.\d+)\.zip$", RegexOptions.Compiled);
        private static readonly Regex s_RegexStrataVersion = new Regex(@"^.*strata\-dist\-(\d+.\d+.\d+)\.zip$", RegexOptions.Compiled);
        #endregion

        #region Properties (Static)
        public static Version DummyVersion => new Version(0, 0, 0);
        #endregion
        
        #region Constructors
        public Manager(String baseDirectory)
        {
            if (String.IsNullOrWhiteSpace(baseDirectory) || !Directory.Exists(baseDirectory))
                throw new ArgumentException("Invalid directory specified.", nameof(baseDirectory));

            m_DirectoryDeployment = Path.Combine(baseDirectory, "Libraries");
            m_DirectoryIKVM = Path.Combine(m_DirectoryDeployment, "IKVM");
            m_DirectoryStrata = Path.Combine(m_DirectoryDeployment, "Strata");
        }
        #endregion

        #region Methods (Static)
        public Boolean CreateDirectories()
        {
            try
            {
                if (Directory.Exists(m_DirectoryIKVM))
                    Directory.Delete(m_DirectoryIKVM, true);

                Directory.CreateDirectory(m_DirectoryIKVM);

                if (Directory.Exists(m_DirectoryStrata))
                    Directory.Delete(m_DirectoryStrata, true);

                Directory.CreateDirectory(m_DirectoryStrata);

                return true;
            }
            catch
            {
                return false;
            }
        }

        public Boolean CreateWrappers()
        {
            FileStream fileStream = null;
            Process process = null;
            RSACryptoServiceProvider provider = null;

            try
            {
                String compilerExecutable = Path.Combine(m_DirectoryIKVM, "ikvmc.exe");

                if (!File.Exists(compilerExecutable))
                    return false;

                String strataVersion = m_VersionStrata.ToString();
                String[] strataWrappers;

                if (Environment.Is64BitOperatingSystem)
                    strataWrappers = new[] {"Strata32.dll", "Strata64.dll"};
                else
                    strataWrappers = new[] {"Strata32.dll"};

                foreach (String strataWrapper in strataWrappers)
                {
                    String strataWrapperBits = strataWrapper.Replace("Strata", String.Empty).Replace(".dll", String.Empty);
                    String strataWrapperPath = Path.Combine(m_DirectoryDeployment, strataWrapper);
                    String strataWrapperSNK = Path.Combine(m_DirectoryStrata, String.Concat("Strata", strataWrapperBits, ".snk"));

                    CspParameters parameters = new CspParameters {KeyNumber = 2};
                    provider = new RSACryptoServiceProvider(2048, parameters);
                    Byte[] array = provider.ExportCspBlob(!provider.PublicOnly);

                    fileStream = new FileStream(strataWrapperSNK, FileMode.Create, FileAccess.Write);
                    fileStream.Write(array, 0, array.Length);
                    fileStream.Flush();
                    fileStream.Close();

                    String commandParameters = String.Join(" ", new List<String>
                    {
                        String.Concat("-out:\"", strataWrapperPath, "\""),
                        String.Concat("-fileversion:", strataVersion),
                        String.Concat("-version:", strataVersion),
                        String.Concat("-keyfile:\"", strataWrapperSNK, "\""),
                        "-target:library",
                        String.Concat("-platform:", ((strataWrapperBits == "32") ? "x86" : "x64")),
                        String.Concat("-recurse:\"", m_DirectoryStrata, "\\*.jar\""),
                        "-classloader:ikvm.runtime.ClassPathAssemblyClassLoader"
                    });

                    process = new Process
                    {
                        StartInfo =
                        {
                            FileName = compilerExecutable,
                            Arguments = commandParameters,
                            CreateNoWindow = true,
                            RedirectStandardError = false,
                            RedirectStandardOutput = false,
                            UseShellExecute = false
                        }
                    };

                    process.Start();
                    process.WaitForExit();

                    File.Delete(strataWrapperSNK);

                    if (process.ExitCode != 0)
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                provider?.Dispose();
                fileStream?.Dispose();
                process?.Dispose();
            }
        }

        public Boolean DeployIKVM()
        {
            FileStream fileStream = null;
            Stream resourceStream = null;
            ZipArchive archive = null;

            try
            {
                String fileName = String.Concat("IKVM-", m_VersionIKVM, ".zip");

                String resourceName = String.Concat("StrataWrapper.", fileName);
                resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);

                if (resourceStream == null)
                    return false;

                String filePath = Path.Combine(m_DirectoryIKVM, fileName);

                fileStream = new FileStream(filePath, FileMode.Create);
                resourceStream.CopyTo(fileStream);
                fileStream.Flush();
                fileStream.Close();

                archive = ZipFile.OpenRead(filePath);

                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.Contains("/bin/"))
                        entry.ExtractToFile(Path.Combine(m_DirectoryIKVM, entry.Name));
                }

                archive.Dispose();
                archive = null;

                File.Delete(filePath);

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                fileStream?.Dispose();
                resourceStream?.Dispose();
                archive?.Dispose();
            }
        }

        public Boolean DeployStrata()
        {
            HttpClient client = null;
            HttpContent content = null;
            HttpResponseMessage response = null;

            FileStream fileStream = null;
            Stream responseStream = null;
            ZipArchive archive = null;

            try
            {
                String versionText = m_VersionStrata.ToString();
                String zipPath = Path.Combine(m_DirectoryStrata, String.Concat("Strata-", versionText, ".zip"));
                Uri uri = new Uri(String.Concat("https://github.com/OpenGamma/Strata/releases/download/v", versionText, "/strata-dist-", versionText, ".zip"));

                client = new HttpClient();
                client.DefaultRequestHeaders.Accept.ParseAdd("application/zip");
                client.DefaultRequestHeaders.UserAgent.ParseAdd("StrataWrapper");

                response = client.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead).Result;
                response.EnsureSuccessStatusCode();

                content = response.Content;
                responseStream = content.ReadAsStreamAsync().Result;

                fileStream = new FileStream(zipPath, FileMode.Create);
                responseStream.CopyTo(fileStream);
                fileStream.Flush();
                fileStream.Close();

                archive = ZipFile.OpenRead(zipPath);

                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.Contains("/lib/") && entry.Name.EndsWith(".jar", StringComparison.OrdinalIgnoreCase) && !entry.Name.Contains("examples"))
                        entry.ExtractToFile(Path.Combine(m_DirectoryStrata, entry.Name));
                }

                archive.Dispose();
                archive = null;

                File.Delete(zipPath);

                foreach (String jarPath in Directory.GetFiles(m_DirectoryStrata, "*.jar"))
                {
                    archive = ZipFile.Open(jarPath, ZipArchiveMode.Update);

                    foreach (ZipArchiveEntry entry in archive.Entries.ToList())
                    {
                        if ((entry.Name == "module-info.class") || (entry.Name == "package-info.class"))
                            entry.Delete();
                    }

                    archive.Dispose();
                    archive = null;
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                client?.Dispose();
                content?.Dispose();
                response?.Dispose();

                fileStream?.Dispose();
                responseStream?.Dispose();
                archive?.Dispose();
            }
        }

        public Boolean FinalizeWrappers()
        {
            try
            {
                if (!Directory.Exists(m_DirectoryIKVM))
                    return false;

                foreach (String filePath in Directory.GetFiles(m_DirectoryIKVM, "IKVM.*.dll", SearchOption.TopDirectoryOnly))
                {
                    String fileName = Path.GetFileName(filePath);

                    if (String.IsNullOrWhiteSpace(fileName))
                        continue;

                    File.Copy(filePath, Path.Combine(m_DirectoryDeployment, fileName));
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public Boolean PerformCleanup()
        {
            try
            {
                if (!Directory.Exists(m_DirectoryDeployment))
                    return true;

                foreach (String filePath in Directory.GetFiles(m_DirectoryDeployment, "*.dll", SearchOption.TopDirectoryOnly))
                {
                    String fileName = Path.GetFileName(filePath);

                    if (String.IsNullOrWhiteSpace(fileName))
                        continue;

                    if (!fileName.StartsWith("IKVM.", StringComparison.OrdinalIgnoreCase) && !fileName.StartsWith("Strata", StringComparison.OrdinalIgnoreCase))
                        continue;

                    File.Delete(filePath);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        public Version GetVersionIKVM()
        {
            m_VersionIKVM = DummyVersion;

            foreach (String resourceName in Assembly.GetExecutingAssembly().GetManifestResourceNames())
            {
                Match match = s_RegexCompilerVersion.Match(resourceName);

                if (!match.Success)
                    continue;

                Version version = Version.Parse(match.Groups[1].Value);

                if (version <= m_VersionIKVM)
                    continue;

                m_VersionIKVM = version;
            }

            return m_VersionIKVM;
        }

        public Version GetVersionStrata()
        {
            Uri uri = new Uri("https://api.github.com/repos/OpenGamma/Strata/releases");

            HttpClient client = null;
            HttpResponseMessage response = null;
            MemoryStream stream = null;

            m_VersionStrata = DummyVersion;

            try
            {
                client = new HttpClient();
                client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github.v3.text+json");
                client.DefaultRequestHeaders.UserAgent.ParseAdd("StrataWrapper");

                response = client.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead).Result;
                response.EnsureSuccessStatusCode();

                stream = new MemoryStream(response.Content.ReadAsByteArrayAsync().Result);

                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(List<GitHubRelease>));
                List<GitHubRelease> releases = (List<GitHubRelease>)serializer.ReadObject(stream);

                foreach (GitHubRelease release in releases)
                {
                    if (release.IsPreRelease)
                        continue;

                    String versionText = null;

                    foreach (GitHubReleaseAsset asset in release.Assets)
                    {
                        if ((asset.ContentType != "application/zip") || String.IsNullOrWhiteSpace(asset.Name) || String.IsNullOrWhiteSpace(asset.Url))
                            continue;

                        Match match = s_RegexStrataVersion.Match(asset.Url);

                        if (!match.Success)
                            continue;

                        versionText = match.Groups[1].Value;

                        break;
                    }

                    if (versionText == null)
                        continue;

                    Version version = Version.Parse(versionText);

                    if (version <= m_VersionStrata)
                        continue;

                    m_VersionStrata = version;
                }

                return m_VersionStrata;
            }
            catch
            {
                m_VersionStrata = null;
            }
            finally
            {
                client?.Dispose();
                response?.Dispose();
                stream?.Dispose();
            }

            return m_VersionStrata;
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(m_DirectoryIKVM, true);
            }
            catch { }

            try
            {
                Directory.Delete(m_DirectoryStrata, true);
            }
            catch { }
        }
        #endregion

        #region Nesting (Classes)
        [DataContract]
        private sealed class GitHubRelease
        {
            #region Properties
            [DataMember(Name = "id")]
            public UInt32 Id { get; private set; }

            [DataMember(Name = "tag_name")]
            public String Tag { get; private set; }

            [DataMember(Name = "name")]
            public String Name { get; private set; }

            [DataMember(Name = "prerelease")]
            public Boolean IsPreRelease { get; private set; }

            [DataMember(Name = "assets")]
            public IList<GitHubReleaseAsset> Assets { get; private set; }
            #endregion
        }

        [DataContract]
        private sealed class GitHubReleaseAsset
        {
            #region Properties
            [DataMember(Name = "id")]
            public UInt32 Id { get; private set; }

            [DataMember(Name = "name")]
            public String Name { get; private set; }

            [DataMember(Name = "state")]
            public String State { get; private set; }

            [DataMember(Name = "content_type")]
            public String ContentType { get; private set; }

            [DataMember(Name = "size")]
            public UInt32 Size { get; private set; }

            [DataMember(Name = "browser_download_url")]
            public String Url { get; private set; }
            #endregion
        }
        #endregion
    }
}