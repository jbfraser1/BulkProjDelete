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
         = "Deletes Projects from a Project Server.\n\n"
            + "Usage: DeleteProjects -url http[s]://PWAServer/pwa/ -inputfile path\\filename\n"
            + "       [-deletewsssites] [-deletearchived] [-wait] [-verify]\n\n"
            + "Options:\n"
            + "   -url pwaurl      Specify the url for the PWA instance on which to delete\n"
            + "                       sites. Required.\n"
            + "   -inputfile path  Specify a text file listing projects to be deleted.\n"
            + "                       Each project should be on a separate line. Required.\n"
            + "   -deletewsssites  WSS sites related to the deleted projects will be\n"
            + "                       deleted as well. Ignored if -deletearchived is used.\n"
            + "   -deletearchived  The projects are deleted from the archive database. If\n"
            + "                       not present, projects are deleted from the draft and\n"
            + "                       published databases.\n"
            + "   -wait            Execution will pause until Project Server processes\n"
            + "                      each job.\n"
            + "   -verify          Command will not actually delete projects or WSS sites.\n\n"
            + "Example:\n"
            + "   deleteprojects -url https://server/pwa/ -file c:\\temp\\oldprojects.txt\n"
            + "         -deletewsssites\n\n";


        static string projectServerUrl = ""; // will include a trailing slash and should have pwa instance.
        static string inputFilePath = ""; //full path including filename.
        static bool deleteWssSites = false;  // true if parameter is set.
        static bool deleteArchived = false;  // true if parameter is set
        static bool wait = false;    //true if parameter is set.
        static bool verify = false;   //true if parameter is set.


        static ProjectWebSvc.ProjectDataSet allProjects = null;
        static ArchiveWebSvc.ArchivedProjectsDataSet allArchiveProjects = null;
        static List<string> ProjectNames = new List<string>();
        static List<Guid> ProjectGuids = new List<Guid>();
        static List<Guid> ProjectVersionGuids = new List<Guid>();


        const string PROJECT_SERVICE_PATH = "_vti_bin/psi/project.asmx";
        const string QUEUESYSTEM_SERVICE_PATH = "_vti_bin/psi/queuesystem.asmx";
        const string PROJECT_ARCHIVE_PATH = "_vti_bin/psi/Archive.asmx";


        [STAThread]
        static void Main(string[] args)
        {

            //validate arguments
            #region Parse Arguments
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
            #endregion

            if (!validArgs)
            {
                Console.WriteLine(usageHelp);
            }
            else
            {

                try
                {

                    #region Web Service Setup
                    //ProjectServerUrl = "http://servername/pwa/";
                    Guid jobId;

                    // Set up the Web service objects
                    ProjectWebSvc.Project projectSvc = new ProjectWebSvc.Project();
                    projectSvc.Url = projectServerUrl + PROJECT_SERVICE_PATH;
                    projectSvc.UseDefaultCredentials = true;

                    QueueSystemWebSvc.QueueSystem q = new QueueSystemWebSvc.QueueSystem();
                    q.Url = projectServerUrl + QUEUESYSTEM_SERVICE_PATH;
                    q.UseDefaultCredentials = true;

                    ArchiveWebSvc.Archive archiveSvc = new ArchiveWebSvc.Archive();
                    archiveSvc.Url = projectServerUrl + PROJECT_ARCHIVE_PATH;
                    archiveSvc.UseDefaultCredentials = true;

                    #endregion

                    #region Read Project List from Server
                    Console.WriteLine("Connecting to Project Server to retrieve project list...");

                    // Read all the projects on the server
                    
                    if (!deleteArchived)
                    {
                        // was allProjects = projectSvc.ReadProjectList();
                        allProjects = projectSvc.ReadProjectStatus(
                                     Guid.Empty,
                                     ProjectWebSvc.DataStoreEnum.WorkingStore,
                                     string.Empty,
                                     (int)PSLibrary.Project.ProjectType.Project);

                    } else
                    {
                        allArchiveProjects = archiveSvc.ReadArchivedProjectsList();
                    }
                    #endregion


                    #region Read text input file and find projects
                    Console.WriteLine("Reading input file...");

                    StreamReader SR;
                    int inputLines = 0;
                    int projectsNotFound = 0;
                    string projName;
                    SR = File.OpenText(inputFilePath);
                    projName = SR.ReadLine().Trim();

                    while (projName != null)
                    {
                        bool foundProject = false;
                        inputLines++;

                        if (!deleteArchived)
                        {
                            // loop through the dataset looking for a matching project.
                            foreach (DataRow projectRow in allProjects.Project)
                            {
                                if (((String)projectRow[allProjects.Project.PROJ_NAMEColumn]).ToLower()
                                    .Equals(projName.ToLower()))
                                {
                                    foundProject = true;
                                    ProjectNames.Add((String)projectRow[allProjects.Project.PROJ_NAMEColumn]);
                                    ProjectGuids.Add((Guid)projectRow[allProjects.Project.PROJ_UIDColumn]);
                                    break;
                                }
                            }
                        }
                        else
                        {
                            foreach (DataRow projectRow in allArchiveProjects.Projects)
                            {
                                if (((String)projectRow[allArchiveProjects.Projects.PROJ_NAMEColumn]).ToLower()
                                    .Equals(projName.ToLower()))
                                {
                                    foundProject = true;
                                    ProjectNames.Add((String)projectRow[allArchiveProjects.Projects.PROJ_NAMEColumn]);
                                    ProjectGuids.Add((Guid)projectRow[allArchiveProjects.Projects.PROJ_UIDColumn]);
                                    ProjectVersionGuids.Add(
                                        (Guid)projectRow[allArchiveProjects.Projects.PROJ_VERSION_UIDColumn]);

                                    // Don't break: There might be multiple rows that match this project, for multiple versions
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

                        projName = SR.ReadLine();
                    }
                    SR.Close();
                    #endregion

                    #region Confirm Project deletion
                    //Console.WriteLine(inputLines.ToString() + " lines were read from input file.";
                    Console.WriteLine(projectsNotFound.ToString() + " projects listed in the input file were not found.");
                    if (!deleteArchived)
                    {
                        Console.WriteLine(ProjectNames.Count.ToString() + " projects will be deleted from draft and published dbs");
                        Console.WriteLine("   (out of " + allProjects.Project.Count.ToString() + " projects on the server.)");

                    }
                    else
                    {
                        Console.WriteLine(ProjectNames.Count.ToString() + " projects will be deleted from archive db");
                        Console.WriteLine("   (out of " + allArchiveProjects.Projects.Count.ToString() + " projects in the Archive db.)");
                        Console.WriteLine(" Note that a project may be counted multiple times: multiple versions may\n be present in the archive db.");
                    }

                    if (!verify && ProjectNames.Count > 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("WARNING! About to delete projects. This action is not reversible!");
                        Console.ResetColor();
                        Console.Write("Enter Y if you want to continue: ");
                        string confirm = Console.ReadLine();
                        if (!("y".Equals(confirm.ToLower())))
                        {
                            verify = true;
                        }
                    }

                    #endregion


                    #region Delete Project

                    if (!verify && ProjectNames.Count > 0)
                    {
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
                            Console.WriteLine("Queued delete for: " + ProjectNames[i]);
                            if (wait)
                            {
                                WaitForQueue(q, jobId);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("No deletes. Verifying only.");
                    }
                    #endregion

                    Console.WriteLine("Execution Complete.");
                }
                #region Exceptions and Final
                // This region is from MSDN, but with minor modifications to the messages:
                // http://msdn.microsoft.com/en-us/library/websvcproject.project.queuedeleteprojects.aspx

                catch (SoapException ex)
                {
                    PSLibrary.PSClientError error = new PSLibrary.PSClientError(ex);
                    PSLibrary.PSErrorInfo[] errors = error.GetAllErrors();
                    string errMess = "==============================\r\nError: \r\n";
                    for (int i = 0; i < errors.Length; i++)
                    {
                        errMess += "\n" + ex.Message.ToString() + "\r\n";
                        errMess += "".PadRight(30, '=') + "\r\nPSCLientError Output:\r\n \r\n";
                        errMess += errors[i].ErrId.ToString() + "\n";

                        for (int j = 0; j < errors[i].ErrorAttributes.Length; j++)
                        {
                            errMess += "\r\n\t" + errors[i].ErrorAttributeNames()[j] + ": " + errors[i].ErrorAttributes[j];
                        }
                        errMess += "\r\n".PadRight(30, '=');
                    }
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(errMess);
                }
                catch (WebException ex)
                {
                    string errMess = ex.Message.ToString() +
                       "\n\nError connecting to Project Server. Please check the URL, your permissions\n"
                    + "to connect to the server, and the Project Server Queuing Service.\n";
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

            // may want to get rid of this for automation?
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();

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
                        Console.WriteLine("Queue Job failed.");
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.WriteLine("Job State: " + jobState + " for Job ID: " + jobId);
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