// ExportLimbusService.cs
using System;
using System.IO;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using System.Text.Json;

namespace ExportLimbusService
{
    public partial class LimbusExportService : ServiceBase
    {
        // Configuration
        private string _exportFolder;
        private string _dicomFolder;
        private string _archiveFolder;
        private string _importFolder;
        private Regex _fileNameRegex;
        private int _archiveRetentionDays;
        private bool _enableImportCleanup;
        private readonly string _configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");

        private FileSystemWatcher _watcher;
        private EventLog _eventLog;

        // === Champs pour garder le service actif ===
        private Thread _workerThread;
        private ManualResetEvent _stopEvent;

        public LimbusExportService()
        {
            ServiceName = "LimbusExport";
            _eventLog = new EventLog();

            if (!EventLog.SourceExists("LimbusExportSource"))
            {
                EventLog.CreateEventSource("LimbusExportSource", "LimbusExportLog");
            }

            _eventLog.Source = "LimbusExportSource";
            _eventLog.Log = "LimbusExportLog";
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                base.OnStart(args);
                LogInfo("Service LimbusExport démarré");

                // Chargement de la configuration
                LoadConfiguration();

                // Vérifie que les dossiers existent
                EnsureDirectoryExists(_exportFolder);
                EnsureDirectoryExists(_dicomFolder);
                EnsureDirectoryExists(_archiveFolder);
                if (_enableImportCleanup && !string.IsNullOrEmpty(_importFolder))
                {
                    EnsureDirectoryExists(_importFolder);
                }

                // Nettoyage au démarrage (une seule fois)
                PerformStartupCleanup();

                // Traitement des fichiers existants au démarrage
                ProcessExistingFiles();

                // Configure le FileSystemWatcher
                _watcher = new FileSystemWatcher
                {
                    Path = _exportFolder,
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime,
                    Filter = "*.dcm",
                    EnableRaisingEvents = true
                };

                // Abonnement aux événements
                _watcher.Created += OnFileCreated;

                LogInfo($"Surveillance du dossier {_exportFolder} en cours");
            }
            catch (Exception ex)
            {
                LogError($"Erreur lors du démarrage du service: {ex}");
                Stop();
                return;
            }

            // === Initialisation du thread bloquant pour garder le service actif ===
            _stopEvent = new ManualResetEvent(false);
            _workerThread = new Thread(() => _stopEvent.WaitOne())
            {
                IsBackground = true
            };
            _workerThread.Start();
        }

        private void PerformStartupCleanup()
        {
            try
            {
                LogInfo("Début du processus de nettoyage au démarrage");

                // Nettoyage du dossier d'archive (fichiers de plus de X jours)
                CleanupArchiveFolder();

                // Nettoyage du dossier IMPORT (fichiers non-.dcm)
                if (_enableImportCleanup && !string.IsNullOrEmpty(_importFolder))
                {
                    CleanupImportFolder();
                }

                LogInfo("Processus de nettoyage au démarrage terminé");
            }
            catch (Exception ex)
            {
                LogError($"Erreur lors du processus de nettoyage au démarrage: {ex.Message}");
            }
        }

        private void CleanupArchiveFolder()
        {
            try
            {
                if (!Directory.Exists(_archiveFolder))
                {
                    LogWarning($"Le dossier d'archive n'existe pas: {_archiveFolder}");
                    return;
                }

                DateTime cutoffDate = DateTime.Now.AddDays(-_archiveRetentionDays);
                string[] files = Directory.GetFiles(_archiveFolder);
                int deletedCount = 0;

                foreach (string file in files)
                {
                    try
                    {
                        FileInfo fileInfo = new FileInfo(file);
                        if (fileInfo.CreationTime < cutoffDate)
                        {
                            File.Delete(file);
                            deletedCount++;
                            LogInfo($"Fichier supprimé de l'archive: {Path.GetFileName(file)} (créé le {fileInfo.CreationTime:yyyy-MM-dd})");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogError($"Erreur lors de la suppression du fichier {file}: {ex.Message}");
                    }
                }

                LogInfo($"Nettoyage de l'archive terminé: {deletedCount} fichiers supprimés (plus anciens que {_archiveRetentionDays} jours)");
            }
            catch (Exception ex)
            {
                LogError($"Erreur lors du nettoyage du dossier d'archive: {ex.Message}");
            }
        }

        private void CleanupImportFolder()
        {
            try
            {
                if (!Directory.Exists(_importFolder))
                {
                    LogWarning($"Le dossier IMPORT n'existe pas: {_importFolder}");
                    return;
                }

                string[] files = Directory.GetFiles(_importFolder);
                int deletedCount = 0;

                foreach (string file in files)
                {
                    try
                    {
                        string extension = Path.GetExtension(file).ToLower();
                        if (extension != ".dcm")
                        {
                            File.Delete(file);
                            deletedCount++;
                            LogInfo($"Fichier non-DICOM supprimé du dossier IMPORT: {Path.GetFileName(file)}");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogError($"Erreur lors de la suppression du fichier {file}: {ex.Message}");
                    }
                }

                LogInfo($"Nettoyage du dossier IMPORT terminé: {deletedCount} fichiers non-DICOM supprimés");
            }
            catch (Exception ex)
            {
                LogError($"Erreur lors du nettoyage du dossier IMPORT: {ex.Message}");
            }
        }

        private void LoadConfiguration()
        {
            try
            {
                if (!File.Exists(_configPath))
                {
                    // Valeurs par défaut si le fichier de configuration n'existe pas
                    var config = new AppConfig
                    {
                        ExportFolder = @"C:\Users\utilisateur\Desktop\EXPORT",
                        DicomFolder = @"\\chemin\vers\DICOM",
                        ArchiveFolder = @"C:\Users\utilisateur\Desktop\IMPORT\ARCHIVE",
                        ImportFolder = @"C:\Users\utilisateur\Desktop\IMPORT",
                        FileNameRegexPattern = @"^limbus_[^_]+_([A-Za-z0-9]+)(?:_[A-Za-z0-9]+)*\.dcm$",
                        ArchiveRetentionDays = 30,
                        EnableImportCleanup = true
                    };

                    string jsonConfig = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(_configPath, jsonConfig);

                    LogInfo($"Fichier de configuration créé avec les valeurs par défaut: {_configPath}");
                }
                else
                {
                    // Lecture de la configuration existante
                    string jsonConfig = File.ReadAllText(_configPath);
                    var config = JsonSerializer.Deserialize<AppConfig>(jsonConfig);

                    _exportFolder = config.ExportFolder;
                    _dicomFolder = config.DicomFolder;

                    // Mise à jour progressive de la configuration si de nouveaux champs sont ajoutés
                    bool configUpdated = false;

                    if (string.IsNullOrEmpty(config.ArchiveFolder))
                    {
                        config.ArchiveFolder = Path.Combine(Path.GetDirectoryName(_exportFolder), "IMPORT", "ARCHIVE");
                        configUpdated = true;
                        LogInfo($"Ajout du dossier d'archive à la configuration: {config.ArchiveFolder}");
                    }
                    _archiveFolder = config.ArchiveFolder;

                    if (string.IsNullOrEmpty(config.ImportFolder))
                    {
                        config.ImportFolder = Path.Combine(Path.GetDirectoryName(_exportFolder), "IMPORT");
                        configUpdated = true;
                        LogInfo($"Ajout du dossier IMPORT à la configuration: {config.ImportFolder}");
                    }
                    _importFolder = config.ImportFolder;

                    if (string.IsNullOrEmpty(config.FileNameRegexPattern))
                    {
                        config.FileNameRegexPattern = @"^limbus_[^_]+_([A-Za-z0-9]+)(?:_[A-Za-z0-9]+)*\.dcm$";
                        configUpdated = true;
                        LogInfo("Ajout du pattern regex à la configuration");
                    }

                    if (config.ArchiveRetentionDays == 0)
                    {
                        config.ArchiveRetentionDays = 30;
                        configUpdated = true;
                        LogInfo("Ajout de la rétention d'archive à la configuration (30 jours)");
                    }
                    _archiveRetentionDays = config.ArchiveRetentionDays;

                    // EnableImportCleanup par défaut à true si non défini
                    _enableImportCleanup = config.EnableImportCleanup;

                    // Sauvegarde de la configuration mise à jour si nécessaire
                    if (configUpdated)
                    {
                        string updatedJsonConfig = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                        File.WriteAllText(_configPath, updatedJsonConfig);
                        LogInfo("Configuration mise à jour et sauvegardée");
                    }

                    LogInfo($"Configuration chargée depuis: {_configPath}");
                }

                // Compilation du regex à partir de la configuration
                try
                {
                    var config = JsonSerializer.Deserialize<AppConfig>(File.ReadAllText(_configPath));
                    _fileNameRegex = new Regex(config.FileNameRegexPattern, RegexOptions.Compiled);
                    LogInfo($"Regex compilé: {config.FileNameRegexPattern}");
                }
                catch (Exception ex)
                {
                    LogError($"Erreur lors de la compilation du regex: {ex.Message}. Utilisation du regex par défaut.");
                    _fileNameRegex = new Regex(@"^limbus_[^_]+_([A-Za-z0-9]+)(?:_[A-Za-z0-9]+)*\.dcm$", RegexOptions.Compiled);
                }

                LogInfo($"Dossier d'export: {_exportFolder}");
                LogInfo($"Dossier DICOM: {_dicomFolder}");
                LogInfo($"Dossier d'archive: {_archiveFolder}");
                LogInfo($"Dossier IMPORT: {_importFolder}");
                LogInfo($"Rétention archive: {_archiveRetentionDays} jours");
                LogInfo($"Nettoyage IMPORT activé: {_enableImportCleanup}");
            }
            catch (Exception ex)
            {
                LogError($"Erreur lors du chargement de la configuration: {ex.Message}");
                throw;
            }
        }

        protected override void OnStop()
        {
            if (_watcher != null)
            {
                _watcher.Created -= OnFileCreated;
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null;
            }

            // === Signal d'arrêt et attente de fin du thread ===
            _stopEvent.Set();
            if (!_workerThread.Join(3000))
            {
                // en dernier recours
                _workerThread.Abort();
            }

            LogInfo("Service LimbusExport arrêté");
        }

        private void ProcessExistingFiles()
        {
            try
            {
                LogInfo("Traitement des fichiers existants dans le dossier d'export");

                // Récupération de tous les fichiers .dcm du dossier d'export
                string[] existingFiles = Directory.GetFiles(_exportFolder, "*.dcm");
                LogInfo($"{existingFiles.Length} fichiers trouvés dans le dossier d'export");

                foreach (string filePath in existingFiles)
                {
                    try
                    {
                        string fileName = Path.GetFileName(filePath);
                        LogInfo($"Traitement du fichier existant: {fileName}");

                        // Extraire l'ID patient
                        string patientId = ExtractPatientId(fileName);

                        if (string.IsNullOrEmpty(patientId))
                        {
                            LogWarning($"Impossible d'extraire l'ID patient du fichier: {fileName}");
                            continue;
                        }

                        // Rechercher le dossier correspondant
                        string patientFolder = FindPatientFolder(patientId);

                        if (string.IsNullOrEmpty(patientFolder))
                        {
                            LogWarning($"Dossier patient {patientId} non trouvé dans {_dicomFolder}");
                            continue;
                        }

                        // Copier le fichier
                        string destinationPath = Path.Combine(patientFolder, fileName);
                        File.Copy(filePath, destinationPath, true);
                        LogInfo($"Fichier copié avec succès vers: {destinationPath}");

                        // Déplacer le fichier vers le dossier d'archive
                        MoveToArchive(filePath, fileName);
                    }
                    catch (Exception ex)
                    {
                        LogError($"Erreur lors du traitement du fichier existant {filePath}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"Erreur lors du traitement des fichiers existants: {ex.Message}");
            }
        }

        private void OnFileCreated(object sender, FileSystemEventArgs e)
        {
            try
            {
                // Petit délai pour s'assurer que le fichier est complètement écrit
                Thread.Sleep(500);

                string fileName = Path.GetFileName(e.FullPath);
                LogInfo($"Nouveau fichier détecté: {fileName}");

                // Extraire l'ID patient du nom de fichier
                string patientId = ExtractPatientId(fileName);

                if (string.IsNullOrEmpty(patientId))
                {
                    LogWarning($"Impossible d'extraire l'ID patient du fichier: {fileName}");
                    return;
                }

                LogInfo($"ID patient extrait: {patientId}");

                // Rechercher le dossier correspondant
                string patientFolder = FindPatientFolder(patientId);

                if (string.IsNullOrEmpty(patientFolder))
                {
                    LogWarning($"Dossier patient {patientId} non trouvé dans {_dicomFolder}");
                    return;
                }

                // Copier le fichier
                string destinationPath = Path.Combine(patientFolder, fileName);
                File.Copy(e.FullPath, destinationPath, true);
                LogInfo($"Fichier copié avec succès vers: {destinationPath}");

                // Déplacer le fichier vers le dossier d'archive
                MoveToArchive(e.FullPath, fileName);
            }
            catch (Exception ex)
            {
                LogError($"Erreur lors du traitement du fichier {e.FullPath}: {ex.Message}");
            }
        }

        private void MoveToArchive(string sourcePath, string fileName)
        {
            try
            {
                // Assurez-vous que le dossier d'archive existe
                if (!Directory.Exists(_archiveFolder))
                {
                    Directory.CreateDirectory(_archiveFolder);
                    LogInfo($"Dossier d'archive créé: {_archiveFolder}");
                }

                // Construire le chemin de destination dans le dossier d'archive
                string archivePath = Path.Combine(_archiveFolder, fileName);

                // Vérifier si le fichier existe déjà dans l'archive
                if (File.Exists(archivePath))
                {
                    // Ajouter un timestamp pour éviter les doublons
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                    string extension = Path.GetExtension(fileName);
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                    archivePath = Path.Combine(_archiveFolder, $"{fileNameWithoutExt}_{timestamp}{extension}");
                }

                // Déplacer le fichier vers l'archive
                File.Move(sourcePath, archivePath);
                LogInfo($"Fichier déplacé vers l'archive: {archivePath}");
            }
            catch (Exception ex)
            {
                LogError($"Erreur lors du déplacement du fichier vers l'archive: {ex.Message}");
            }
        }

        private string ExtractPatientId(string fileName)
        {
            Match match = _fileNameRegex.Match(fileName);

            if (match.Success && match.Groups.Count > 1)
            {
                return match.Groups[1].Value;
            }

            return null;
        }

        private string FindPatientFolder(string patientId)
        {
            string patientFolder = Path.Combine(_dicomFolder, patientId);

            if (Directory.Exists(patientFolder))
            {
                return patientFolder;
            }

            // Recherche plus approfondie si nécessaire
            foreach (string folder in Directory.GetDirectories(_dicomFolder))
            {
                if (Path.GetFileName(folder) == patientId)
                {
                    return folder;
                }
            }

            return null;
        }

        private void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                try
                {
                    Directory.CreateDirectory(path);
                    LogInfo($"Dossier créé: {path}");
                }
                catch (Exception ex)
                {
                    throw new DirectoryNotFoundException($"Impossible de créer le dossier {path}: {ex.Message}");
                }
            }
        }

        private void LogInfo(string message)
        {
            _eventLog.WriteEntry(message, EventLogEntryType.Information);
        }

        private void LogWarning(string message)
        {
            _eventLog.WriteEntry(message, EventLogEntryType.Warning);
        }

        private void LogError(string message)
        {
            _eventLog.WriteEntry(message, EventLogEntryType.Error);
        }
    }

    // Classe pour la configuration
    public class AppConfig
    {
        public string ExportFolder { get; set; }
        public string DicomFolder { get; set; }
        public string ArchiveFolder { get; set; }
        public string ImportFolder { get; set; }
        public string FileNameRegexPattern { get; set; }
        public int ArchiveRetentionDays { get; set; }
        public bool EnableImportCleanup { get; set; }
    }

}
