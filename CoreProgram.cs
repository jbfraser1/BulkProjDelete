using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Web.Services.Protocols;
using System.Threading;
using System.Data;
using System.IO;

using PSLibrary = Microsoft.Office.Project.Server.Library;

namespace BulkProjectDelete
{
    class CoreProgram
    {
        static String usageHelp
         = "Deletes Projects from a Project Server.\r\n\r\n"
            + "Usage: BulkProjectDelete -url http[s]://PWAServer/pwa/ -inputfile path\\filename\r\n"
            + "       [-deletewsssites] [-deletearchived] [-wait] [-verify]\r\n\r\n"
            + "Options:\r\n"
            + "   -url pwaurl      Url for the PWA instance. Required.\r\n"
            + "   -inputfile path  Specifies a text file listing projects to be deleted.\r\n"
            + "                     Each project should be on a separate line. Required.\r\n"
            + "   -deletewsssites  Deletes project workspaces for specifed projects. Ignored\r\n"
            + "                     if -deletearchived is used.\r\n"
            + "   -deletearchived  Projects are deleted only from the archive database. If not\r\n"
            + "                     present, projects are deleted from draft & published dbs.\r\n"
            + "   -keeplatest      Works with -deletearchived option. The most recent copy of\r\n"
            + "                     a project in the archive db will not be deleted.\r\n"
            + "   -wait            Wait for the server to finish each job.\r\n"
            + "   -verify          No deletions will be performed.\r\n\r\n"
            + "Example:\r\n"
            + "   BulkProjectDelete -url https://server/pwa/ -file c:\\temp\\oldprojects.txt\r\n"
            + "         -deletewsssites\r\n\r\n";


        static string projectServerUrl = ""; // will include a trailing slash and should have pwa instance.
        static string inputFilePath = ""; //full path including filename.
        static bool deleteWssSites = false;  // true if parameter is set.
        static bool deleteArchived = false;  // true if parameter is set
        static bool wait = false;    //true if parameter is set.
        static bool verify = false;   //true if parameter is set.
        static bool keeplatest = false; //true if parameter is set.


        static ProjectWebSvc.ProjectDataSet projectProjects = null;
        static ProjectWebSvc.ProjectDataSet masterProjects = null;
        static ProjectWebSvc.ProjectDataSet lightweightProjects = null;
        static ProjectWebSvc.ProjectDataSet insertedProjects = null;
        static ArchiveWebSvc.ArchivedProjectsDataSet allArchiveProjects = null;
        static List<string> ProjectNames = new List<string>();
        static List<Guid> ProjectGuids = new List<Guid>();
        static List<Guid> ProjectVersionGuids = new List<Guid>();
        static int projectsNotFound = 0;
        static int keepLatestCount = 0; //count of projects that will be kept if keeplatest is selected.


        const string PROJECT_SERVICE_PATH = "_vti_bin/psi/project.asmx";
        const string QUEUESYSTEM_SERVICE_PATH = "_vti_bin/psi/queuesystem.asmx";
        const string PROJECT_ARCHIVE_PATH = "_vti_bin/psi/Archive.asmx";

        static ProjectWebSvc.Project projectSvc;
        static QueueSystemWebSvc.QueueSystem queueSvc;
        static ArchiveWebSvc.Archive archiveSvc;

        [STAThread]
        static void Main(string[] args)
        {

            if (!ParseArgs(args))
            {
                Console.WriteLine(usageHelp);
            }
            else
            {
                try  //need to move these handlers into the methods
                {
                    SetupWebSvc();

                    ReadProjectsFromServer();

                    ReadInputFile();

                    WriteSummary();

                    if (verify)
                    {
                        Console.WriteLine("!!! No deletions performed! Verifying only.");
                    } else if (ProjectNames.Count == 0)
                    {
                        Console.WriteLine("No projects found to delete.");

                    } else if (Confirm())
                    {
                        DeleteProjects();
                    }
                    else
                    {
                        Console.WriteLine("Deletion aborted.");
                    }

                    Console.WriteLine("Execution Complete.");
                }
                #region Error handling and finally
                    // need to move this into the individual methods, handle errors there with more specific messages.
                catch (SoapException ex)
                {
                    PSLibrary.PSClientError error = new PSLibrary.PSClientError(ex);
                    PSLibrary.PSErrorInfo[] errors = error.GetAllErrors();
                    string errMess = "==============================\r\nError: \r\n";
                    foreach (PSLibrary.PSErrorInfo suberror in errors)
                    {
                        errMess += "\r\n" + ex.Message.ToString() + "\r\n\r\n"
                           + "Sub-error:\r\n\r\n"
                           + suberror.ErrId.ToString() + "\r\n";

                        for (int i = 0; i < suberror.ErrorAttributes.Length; i++)
                        {
                            errMess += "\r\n\t" + suberror.ErrorAttributeNames()[i] + ": " + suberror.ErrorAttributes[i];
                        }
                        errMess += "==============================\r\n";
                    }
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(errMess);
                }
                catch (WebException ex)
                {
                    string errMess = ex.Message.ToString() +
                       "\r\n\r\nError connecting to Project Server. Please check the URL, your permissions\r\n"
                    + "to connect to the server, and the Project Server Queuing Service.\r\n";
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: " + errMess);
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: " + ex.Message);
                }
                finally
                {
                    Console.ResetColor();

                }
                #endregion
            }

#if DEBUG
            // may want to get rid of this for automation?
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
#endif

        }

        // Returns true if arguments are valid, false if there's a problem.
        static private bool ParseArgs(string[] args)
        {

            bool validArgs = true;
            bool hasURL = false;
            bool hasFile = false;
            //when done parsing, all three of these need to be true to continue.

            for (int i = 0; i < args.Length; i++)
            {
                string arg = args[i].ToLower();
                if (arg.Equals("help") || arg.Equals("?"))
                {
                    validArgs = false;
                    break;
                }

                arg = arg.Substring(1); // strip the leading dash or slash.

                switch (arg)
                {
                    case "url":
                        i++;
                        // the pwa url must start with "http":
                        if (!(args[i].Substring(0, 4).ToLower().Equals("http")))
                        {
                            validArgs = false;
                        }
                        else
                        {
                            // if there isn't a trailing slash, add one:
                            if (args[i].Substring(
                                  args[i].Length - 1, 1)
                                  .Equals("/"))
                            {
                                projectServerUrl = args[i];
                            }
                            else
                            {
                                projectServerUrl = args[i] + "/";
                            }
                        }
                        hasURL = true;
                        break;
                    case "inputfile":
                        i++;
                        inputFilePath = args[i];
                        hasFile = true;
                        break;
                    case "deletewsssites":
                        deleteWssSites = true;
                        break;
                    case "deletearchived":
                        deleteArchived = true;
                        break;
                    case "keeplatest":
                        keeplatest = true;
                        break;
                    case "wait":
                        wait = true;
                        break;
                    case "verify":
                        verify = true;
                        break;
                    default:
                        validArgs = false;
                        break;
                }

            }

            if (!hasURL || !hasFile)
            {
                validArgs = false;
            }

            return validArgs;
        }

        static private void SetupWebSvc()
        {

            //ProjectServerUrl = "http://servername/pwa/";

            // Set up the Web service objects
            projectSvc = new ProjectWebSvc.Project();
            projectSvc.Url = projectServerUrl + PROJECT_SERVICE_PATH;
            projectSvc.UseDefaultCredentials = true;

            queueSvc = new QueueSystemWebSvc.QueueSystem();
            queueSvc.Url = projectServerUrl + QUEUESYSTEM_SERVICE_PATH;
            queueSvc.UseDefaultCredentials = true;

            archiveSvc = new ArchiveWebSvc.Archive();
            archiveSvc.Url = projectServerUrl + PROJECT_ARCHIVE_PATH;
            archiveSvc.UseDefaultCredentials = true;

        }

        static private void ReadProjectsFromServer()
        {

            Console.WriteLine("Connecting to Project Server to retrieve project list...");

            // Read all the projects on the server

            if (!deleteArchived)
            {
                // was allProjects = projectSvc.ReadProjectList();
                projectProjects = projectSvc.ReadProjectStatus(
                             Guid.Empty,
                             ProjectWebSvc.DataStoreEnum.WorkingStore,
                             string.Empty,
                             (int)PSLibrary.Project.ProjectType.Project);

                // Need to handle three other types of projects:
                masterProjects = projectSvc.ReadProjectStatus(
                         Guid.Empty,
                         ProjectWebSvc.DataStoreEnum.WorkingStore,
                         string.Empty,
                         (int)PSLibrary.Project.ProjectType.MasterProject);

                lightweightProjects = projectSvc.ReadProjectStatus(
                         Guid.Empty,
                         ProjectWebSvc.DataStoreEnum.WorkingStore,
                         string.Empty,
                         (int)PSLibrary.Project.ProjectType.LightWeightProject);

                insertedProjects = projectSvc.ReadProjectStatus(
                         Guid.Empty,
                         ProjectWebSvc.DataStoreEnum.WorkingStore,
                         string.Empty,
                         (int)PSLibrary.Project.ProjectType.InsertedProject);


            }
            else
            {
                allArchiveProjects = archiveSvc.ReadArchivedProjectsList();
            }
        }

        /// <summary>
        /// Reads the input file and matches the projects to the project lists read from the Project Server.
        /// Needs to be broken into smaller methods.
        /// </summary>
        static private void ReadInputFile()
        {
            Console.WriteLine("Reading input file...");

            StreamReader SR;
            int inputLines = 0;

            string projName;
            SR = new StreamReader(inputFilePath, Encoding.Default, true);

            projName = SR.ReadLine();

            while (projName != null)
            {
                projName = projName.Trim();
                if (!("".Equals(projName)))  //skip empty lines.
                {
                    bool foundProject = false;
                    inputLines++;

                    if (!deleteArchived)
                    {
                        // loop through the dataset looking for a matching project.
                        foreach (DataRow projectRow in projectProjects.Project)
                        {
                            if (((String)projectRow[projectProjects.Project.PROJ_NAMEColumn]).ToLower()
                                .Equals(projName.ToLower()))
                            {
                                foundProject = true;
                                ProjectNames.Add((String)projectRow[projectProjects.Project.PROJ_NAMEColumn]);
                                ProjectGuids.Add((Guid)projectRow[projectProjects.Project.PROJ_UIDColumn]);
                                break;
                            }
                        }
                        // masterprojects.
                        if (!foundProject)
                        {
                            foreach (DataRow projectRow in masterProjects.Project)
                            {
                                if (((String)projectRow[masterProjects.Project.PROJ_NAMEColumn]).ToLower()
                                    .Equals(projName.ToLower()))
                                {
                                    foundProject = true;
                                    ProjectNames.Add((String)projectRow[masterProjects.Project.PROJ_NAMEColumn]);
                                    ProjectGuids.Add((Guid)projectRow[masterProjects.Project.PROJ_UIDColumn]);
                                    break;
                                }
                            }
                        }
                        // lightweightprojects.
                        if (!foundProject)
                        {
                            foreach (DataRow projectRow in lightweightProjects.Project)
                            {
                                if (((String)projectRow[lightweightProjects.Project.PROJ_NAMEColumn]).ToLower()
                                    .Equals(projName.ToLower()))
                                {
                                    foundProject = true;
                                    ProjectNames.Add((String)projectRow[lightweightProjects.Project.PROJ_NAMEColumn]);
                                    ProjectGuids.Add((Guid)projectRow[lightweightProjects.Project.PROJ_UIDColumn]);
                                    break;
                                }
                            }
                        }
                        if (!foundProject)
                        {
                            //sub projects
                            foreach (DataRow projectRow in insertedProjects.Project)
                            {
                                if (((String)projectRow[insertedProjects.Project.PROJ_NAMEColumn]).ToLower()
                                    .Equals(projName.ToLower()))
                                {
                                    foundProject = true;
                                    ProjectNames.Add((String)projectRow[insertedProjects.Project.PROJ_NAMEColumn]);
                                    ProjectGuids.Add((Guid)projectRow[insertedProjects.Project.PROJ_UIDColumn]);
                                    break;
                                }
                            }
                        }
                    }
                    else  //deleting archived projects.
                    {
                        foreach (DataRow projectRow in allArchiveProjects.Projects)
                        {
                            if (((String)projectRow[allArchiveProjects.Projects.PROJ_NAMEColumn]).ToLower()
                                .Equals(projName.ToLower()))
                            {
                                
                                if (keeplatest) 
                                {
                                    // need to look through the archive projects for one with the same project UID but a more recent version date.
                                    // if not found, don't delete.
                                    bool OldVersion = false; // only set to true once we know that there is a newer project
                                    DateTime DateInQuestion = (DateTime)projectRow[allArchiveProjects.Projects.PROJ_VERSION_DATEColumn];

                                    // the Project Version uid stores a UID that will be shared by multiple versions of the same project. Seems to be a bug in Projects Server.
                                 // http://www.dotnetzone.gr/cs/blogs/pkanavos/archive/2009/02/02/project-server-curios-the-queuedeletearchivedproject-method.aspx

                                    Guid GuidInQuestion = (Guid)projectRow[allArchiveProjects.Projects.PROJ_VERSION_UIDColumn];

                                    foreach (DataRow maybeNewerProject in allArchiveProjects.Projects)
                                    {
                                        if (GuidInQuestion.Equals(
                                            (Guid)maybeNewerProject[allArchiveProjects.Projects.PROJ_VERSION_UIDColumn]))
                                        {
                                            DateTime maybeNewerDate = (DateTime)maybeNewerProject[allArchiveProjects.Projects.PROJ_VERSION_DATEColumn];

                                            if (DateTime.Compare(DateInQuestion, maybeNewerDate) < 0)
                                            {
                                                OldVersion = true;
                                                foundProject = true;
                                                ProjectNames.Add((String)projectRow[allArchiveProjects.Projects.PROJ_NAMEColumn]);
                                                ProjectGuids.Add((Guid)projectRow[allArchiveProjects.Projects.PROJ_UIDColumn]);
                                                ProjectVersionGuids.Add(
                                                      (Guid)projectRow[allArchiveProjects.Projects.PROJ_VERSION_UIDColumn]);
                                                break;
                                            }
                                        }
                                    }

                                    if (!OldVersion)
                                    {
                                        // increase count of projects that are the latest version.
                                        keepLatestCount++;
                                    }
                                } else
                                {
                                    // not keeping latest, don't need to compare to other projects in the archive set.

                                    foundProject = true;

                                    ProjectNames.Add((String)projectRow[allArchiveProjects.Projects.PROJ_NAMEColumn]);
                                    ProjectGuids.Add((Guid)projectRow[allArchiveProjects.Projects.PROJ_UIDColumn]);
                                    ProjectVersionGuids.Add(
                                        (Guid)projectRow[allArchiveProjects.Projects.PROJ_VERSION_UIDColumn]);

                                    // Don't break: There might be multiple rows that match this project, for multiple versions
                                }
                            }
                        }
                    }

                    if (foundProject)
                    {
                        Console.WriteLine("Found: " + projName);
                    }
                    else
                    {
                        projectsNotFound++;
                        Console.WriteLine("FAILED to find: " + projName);
                    }
                }
                projName = SR.ReadLine();
            }
            SR.Close();
        }

        static private void WriteSummary()
        {
           
            //Console.WriteLine(inputLines.ToString() + " lines were read from input file.";
            if (projectsNotFound > 0)
            {
                Console.WriteLine(projectsNotFound.ToString() + " projects listed in the input file were not found.");
            }
            else
            {
                Console.WriteLine("All projects listed in input file were found.");
            }

            if (!deleteArchived)
            {
                int projectsOnServer = projectProjects.Project.Count +
                    masterProjects.Project.Count +
                    lightweightProjects.Project.Count +
                    insertedProjects.Project.Count;
                Console.WriteLine(ProjectNames.Count.ToString() + " projects will be deleted from draft and published dbs");
                Console.WriteLine("   (out of " + projectsOnServer.ToString() + " projects on the server.)");

            }
            else
            {
                Console.WriteLine(ProjectNames.Count.ToString() + " projects will be deleted from archive db");
                Console.WriteLine("   (out of " + allArchiveProjects.Projects.Count.ToString() + " projects in the archive db.)");
                if (keeplatest)
                {
                    Console.WriteLine("   " + keepLatestCount.ToString() + " projects were found to be the latest version. They will\n   be kept.");
                }
                Console.WriteLine(" Note that a project may be counted multiple times: multiple versions may\n   be present in the archive db.");
            }
        }

        static private bool Confirm()
        {
            // Confirm Project deletion
            //return true if want to continue.
            
            if (!verify && ProjectNames.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("WARNING! About to delete projects. This action is not reversible!");
                Console.ResetColor();
                Console.Write("Enter Y if you want to continue: ");
                string confirm = Console.ReadLine();
                if ("y".Equals(confirm.ToLower()))
                {
                    return true;
                }
            }

            return false;
        }

        static private void DeleteProjects()
        {

                Guid jobId;
                Console.WriteLine("Beginning to delete projects...");

                Guid[] deleteGuid = new Guid[1];
                for (int i = 0; i < ProjectGuids.Count; i++)
                {

                    jobId = Guid.NewGuid();
                    if (!deleteArchived)
                    {
                        deleteGuid[0] = ProjectGuids[i];
                        projectSvc.QueueDeleteProjects(jobId, deleteWssSites, deleteGuid, true);
                    }
                    else
                    {
                        // Note the bug below: the Guids must be handed in the wrong positions!
                        // If they are reversed, the queue reports success, but nothing is deleted.

                        archiveSvc.QueueDeleteArchivedProject(
                             jobId,
                             ProjectVersionGuids[i],
                             ProjectGuids[i]);

                    }
                    Console.WriteLine("Queued delete for: " +
                        "(" + (i+1).ToString() + " of " + ProjectGuids.Count.ToString() +") "
                        + ProjectNames[i]);
                    if (wait)
                    {
                        WaitForQueue(queueSvc, jobId);
                    }
                }
            }

        static private void WaitForQueue(QueueSystemWebSvc.QueueSystem q, Guid jobId)
        {

            // This method is from MSDN, but I've added minor modifications to the wait time handling.
            // http://msdn.microsoft.com/en-us/library/websvcproject.project.queuedeleteprojects.aspx

            QueueSystemWebSvc.JobState jobState;
            const int QUEUE_WAIT_TIME = 10; // ten seconds
            bool jobDone = false;
            string xmlError = string.Empty;
            int wait = 0;
            int loop = 0;

            //Wait for the project to get through the queue
            // - Get the estimated wait time in seconds
            wait = q.GetJobWaitTime(jobId);

            // - Wait for it
            wait = Math.Min(wait, QUEUE_WAIT_TIME);
            Thread.Sleep(wait * 1000);
            // - Wait until it is done.

            do
            {
                // - Get the job state
                jobState = q.GetJobCompletionState(jobId, out xmlError);

                if (jobState == QueueSystemWebSvc.JobState.Success)
                {
                    jobDone = true;
                }
                else
                {
                    if (jobState == QueueSystemWebSvc.JobState.Unknown
                    || jobState == QueueSystemWebSvc.JobState.Failed
                    || jobState == QueueSystemWebSvc.JobState.FailedNotBlocking
                    || jobState == QueueSystemWebSvc.JobState.CorrelationBlocked
                    || jobState == QueueSystemWebSvc.JobState.Canceled)
                    {
                        // Used to throw exception, but now just displaying error.
                        //throw (new ApplicationException("Queue request " + jobState + " for Job ID " + jobId + ".\r\n" + xmlError));
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine("Queue Job failed. See the server Manage Queue interface for more\r\ninformation.");
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.WriteLine("Status: " + jobState + "   (Job ID: " + jobId + ")");
                        loop++;
                        if (loop >= 10)
                        {
                            loop = 0;
                            Console.WriteLine("Continuing to wait for Project Server Queue...");
                            Console.WriteLine("<CTRL> + C will halt the BulkProjectDelete process.");
                        }
                        wait = q.GetJobWaitTime(jobId);

                        // - Wait for it
                        wait = Math.Min(wait, QUEUE_WAIT_TIME);
                        Thread.Sleep(wait * 1000);
                    }
                }
            }
            while (!jobDone);
}

    }
}