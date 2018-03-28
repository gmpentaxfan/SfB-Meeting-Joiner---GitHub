using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
using System.IO;
using System.Net;
using System.DirectoryServices.AccountManagement;
using System.Diagnostics;
using System.Threading;


using Microsoft.Lync.Model; //Updated to Lync 2013 SDK
using Microsoft.Lync.Model.Extensibility;
using Microsoft.Lync.Model.Conversation;
using Microsoft.Lync.Model.Conversation.Sharing;
using Microsoft.Lync.Model.Conversation.AudioVideo; //v1.4

using System.Runtime.InteropServices; //v1.3
using System.Reflection; //v1.5

using System.ComponentModel;
using System.Drawing;

//https://msdn.microsoft.com/en-us/library/office/dn495614(v=exchg.150).aspx

namespace SfB_Meeting_Joiner
{
    public class Main : ApplicationContext
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType); //v1.5

        public static ExchangeService EmailService = new ExchangeService(ExchangeVersion.Exchange2013);
        public static UserPrincipal CurrentEmailAddress = UserPrincipal.Current;
        public static NotifyIcon notifyIcon = new NotifyIcon();
        string icontext = "SfB Meeting Joiner - Right click for more options";
        LyncClient lyncClient;
        
        string DefaultDialHttp = "https://meet.DomainName.ca/";
        string DefaultDomain = "@DomainName.ca";
        int trycount = 0;
        string MeetingSubject;
        public System.Windows.Forms.Timer GetAppointmentTimer = new System.Windows.Forms.Timer();
        public System.Windows.Forms.Timer ConversationTimer = new System.Windows.Forms.Timer();
        double MeetingSleep = 15; //minutes
        public static string conferenceUrl; //v1.5 

        static int RequiredAttendees; //v1.7
        static int OptionalAttendees; //v1.7
        static int TotalAttendees; //v1.7
        static int TotalResources; //v 1.7

        string _PowerPointDeckName = string.Empty;
        string _NativeFileNameAndPath = string.Empty;
        string _NativeFileName = string.Empty;

        static Microsoft.Lync.Model.Conversation.Conversation _conversation;
        ContentSharingModality csmSharing;
        ShareableContent _ShareableContent;

        Process SfBMeetingLauncher = new Process(); //v1.4
        
        public Func<bool> AutoAnswer = () => true; //v1.5

        public Main()
        {
            
            InitLogging(); //v1.5

            SfBMeetingLauncher.StartInfo.WorkingDirectory = AppDomain.CurrentDomain.BaseDirectory; //v1.5
            SfBMeetingLauncher.StartInfo.FileName = "SfB Meeting Launcher.exe";
            
            _PowerPointDeckName = AppDomain.CurrentDomain.BaseDirectory + "help.pptx";

            GetAppointmentTimer.Tick += new EventHandler(AppointmentRecheck);
            ConversationTimer.Tick += new EventHandler(ConversationReCheck);

            notifyIcon.Text = icontext;
            notifyIcon.Icon = SfB_Meeting_Joiner.Properties.Resources.AppIcon;
            notifyIcon.Click += new EventHandler(KeyPressClick); //GM V1.20
            notifyIcon.Visible = true;

            EmailService.UseDefaultCredentials = true;
            string EmailAddress = CurrentEmailAddress.EmailAddress;

            //EmailAddress = "teststudent.achs@student.DomainName.ca"; //Testing

            if (EmailAddress.ToLower().Contains("student.DomainName.ca"))
            {
                EmailService.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                Main.EmailService.UseDefaultCredentials = false;
            }
            else
            {
                try
                {
                    EmailService.AutodiscoverUrl(EmailAddress, RedirectionUrlValidationCallback);
                }
                catch
                {
                    EmailService.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                    //Environment.Exit(0);
                }
            }
            // Get SfB/Lync status
            try
            {
                lyncClient = LyncClient.GetClient();
                //Automation = LyncClient.GetAutomation();
            }
            catch (LyncClientException ex)
            {
                try
                {
                    if (StartLyncProcess())
                        lyncClient = LyncClient.GetClient();
                }
                catch (Exception)
                {
                    //var errorForm = new ErrorForm();
                    //errorForm.ShowDialog();
                    Log.Error("Main: {0}", ex);
                }
            }
            catch
            {
                //MessageBox.Show("Error");
            }

            lyncClient.ConversationManager.ConversationAdded += new EventHandler<ConversationManagerEventArgs>(ConversationManager_ConversationAdded);
            lyncClient.ConversationManager.ConversationRemoved += new EventHandler<ConversationManagerEventArgs>(ConversationManager_ConversationRemoved);
            
            GetAppointmentTimer.Interval = 15000; //15 Seconds
            ConversationTimer.Interval = 15000; //15 Seconds v1.5

            Log.DebugFormat("GetAppointments: Timer Interval '{0}' minutes", GetAppointmentTimer.Interval / 60000);
            GetAppointments(); //v1.1
            //ConversationTimer.Start();
        }

        void KeyPressClick(object sender, EventArgs e)
        {
            //GM V1.00 July 5 2016
            if (Control.ModifierKeys == Keys.Control)
            {
                //MessageBox.Show("Exiting Program");
                notifyIcon.Visible = false;
                Application.Exit();
            }

            if (Control.ModifierKeys == Keys.LShiftKey || Control.ModifierKeys == Keys.RShiftKey)
            {
                //MessageBox.Show("Shift Key");
            }

            if (Control.ModifierKeys == Keys.Alt)
            {
                //MessageBox.Show("ALT Key");
            }
        }

        public void InitLogging() //v1.5
        {
            var logConfig = new FileInfo("Logging.config");
            log4net.Config.XmlConfigurator.ConfigureAndWatch(logConfig);
            Log.Info("SfB Meeting Joiner Starting");
        }

        private void StartContentShare()
        {
            //string selectedItem = (string)Content_Listbox.SelectedItem;
            string selectedItem = "PowerPoint-help.pptx";
            int ShareCount = 0;
           
            foreach (ShareableContent sContent in ((ContentSharingModality)_conversation.Modalities[ModalityTypes.ContentSharing]).ContentCollection)
            {
                ShareCount++;
                Log.DebugFormat("ContentShare_Button: {0}", sContent.Type.ToString() + "-" + sContent.Title);
                
                if (sContent.Type.ToString() + "-" + sContent.Title == selectedItem)
                {
                    int hrReason;
                    try
                    {
                        _ShareableContent = sContent;
                        sContent.Present();
                    }
                    catch
                    {
                        Log.DebugFormat("ContentShare_button Trycount: {0}", trycount);
                        if (trycount >= 3)
                        {
                            break;
                        }

                        if (sContent.CanInvoke(ShareableContentAction.Present, out hrReason))
                        {
                            _ShareableContent = sContent;
                            sContent.Present();
                        }
                        else
                        {
                            trycount ++;
                            //Log.DebugFormat("ContentShare_button Trycount:  {0}", trycount);
                            if (trycount >= 3)
                            {
                                break;
                            }
                            //MessageBox.Show("Cannot present " + sContent.Type.ToString() + ": "+ hrReason.ToString());
                            SharePPT();
                        }
                    }
                }
            }
            if(ShareCount == 0)
            {
                Thread.Sleep(5000);
                SharePPT();
            }
        }

        private bool StartLyncProcess()
        {
            Process lyncClientProcess = null;

            try
            {
                if (File.Exists("C:\\Program Files (x86)\\Microsoft Lync\\communicator.exe"))
                {
                    lyncClientProcess = Process.Start("C:\\Program Files (x86)\\Microsoft Lync\\communicator.exe");
                }
                else if (File.Exists("C:\\Program Files\\Microsoft Lync\\communicator.exe"))
                {
                    lyncClientProcess = Process.Start("C:\\Program Files (x86)\\Microsoft Lync\\communicator.exe");
                }
                // V2 Office 2013
                if (File.Exists("C:\\Program Files (x86)\\Microsoft Office\\Office15\\lync.exe"))
                {
                    lyncClientProcess = Process.Start("C:\\Program Files (x86)\\Microsoft Office\\Office15\\lync.exe");
                }
                else if (File.Exists("C:\\Program Files\\Microsoft Office\\Office15\\lync.exe"))
                {
                    lyncClientProcess = Process.Start("C:\\Program Files (x86)\\Microsoft Office\\Office15\\lync.exe");
                }
                //V2 Office 2016
                if (File.Exists("C:\\Program Files (x86)\\Microsoft Office\\Office15\\lync.exe"))
                {
                    lyncClientProcess = Process.Start("C:\\Program Files (x86)\\Microsoft Office\\Office15\\lync.exe");
                }
                else if (File.Exists("C:\\Program Files\\Microsoft Office\\Office15\\lync.exe"))
                {
                    lyncClientProcess = Process.Start("C:\\Program Files (x86)\\Microsoft Office\\Office15\\lync.exe");
                }

                if (lyncClientProcess != null)
                {
                    System.Threading.Thread.Sleep(2000);
                    return true;
                }
                return false;
            }
            catch (FileNotFoundException ex)
            {
                Log.Error("StartLyncProcess: {0}", ex);
                return false;
            }
        }

        private void GetAppointments() //v1.1
        {
            // Initialize values for the start and end times, and the number of appointments to retrieve.
            try { GetAppointmentTimer.Stop(); }
            catch { };
            GetAppointmentTimer.Interval = 15 * 60000; //15 Minutes
            Log.DebugFormat("GetAppointments: Timer Interval '{0}' minutes", GetAppointmentTimer.Interval/60000);
            GetAppointmentTimer.Start();
            DateTime startDate = DateTime.Now.AddHours(-8);
            DateTime CurrentTime = DateTime.Now;
            DateTime endDate = CurrentTime.AddHours(8);
            const int NUM_APPTS = 16;

            // Initialize the calendar folder object with only the folder ID. 
            CalendarFolder calendar = CalendarFolder.Bind(EmailService, WellKnownFolderName.Calendar, new PropertySet());

            // Set the start and end time and number of appointments to retrieve.
            CalendarView cView = new CalendarView(startDate, endDate, NUM_APPTS);

            // Limit the properties returned to the appointment's subject, start time, and end time.
            cView.PropertySet = new PropertySet(AppointmentSchema.Id, AppointmentSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End);

            // Retrieve a collection of appointments by using the calendar view.
            FindItemsResults<Appointment> appointments = calendar.FindAppointments(cView);
            
            Log.DebugFormat("GetAppointments: The first " + NUM_APPTS + " appointments on your calendar from " + startDate.Date.ToShortDateString() + " " + startDate.TimeOfDay +
                              " to " + endDate.Date.ToShortDateString() + " " + endDate.TimeOfDay + " are:");

            foreach (Appointment a in appointments)
            {
                Log.DebugFormat("GetAppointments: Subject:  {0}", a.Subject.ToString() + " ");
                
                Log.DebugFormat("GetAppointments: Start:  {0}", a.Start.ToString() + " ");
                Log.DebugFormat("GetAppointments: End:  {0}", a.End.ToString());

                
                int StartTime = DateTime.Compare(CurrentTime, a.Start);
                int FinishTime = DateTime.Compare(a.End, CurrentTime);
                
                Appointment appointmentDetailed = Appointment.Bind(EmailService, a.Id, new PropertySet(BasePropertySet.FirstClassProperties) { RequestedBodyType = BodyType.Text });

                //Log.DebugFormat(a.Start + " -- " + appointmentDetailed.JoinOnlineMeetingUrl + " -- " + a.End );
                //Log.DebugFormat(StartTime + " -- " + FinishTime);
                if (appointmentDetailed.JoinOnlineMeetingUrl != null)
                {
                    string tempinfo = "";
                    Log.DebugFormat("GetAppointments: Locations {0}", appointmentDetailed.Location); //v1.7

                    TotalResources = appointmentDetailed.Resources.Count;
                    if(TotalResources == 0)
                    {
                        Array TempLocation = appointmentDetailed.Location.Split(';');
                        foreach(string location in TempLocation)
                        {
                            TotalResources++;
                        }
                        Log.DebugFormat("GetAppointments: Counted Resources {0}", TotalResources); //v1.7
                    }
                    else
                    {
                       
                        foreach (Attendee at in appointmentDetailed.Resources)//v1.7
                        {
                            tempinfo = at.Address + "," + tempinfo;
                        }
                        Log.DebugFormat("GetAppointments: Appointment Resources {0}", appointmentDetailed.Resources.Count); //v1.7
                        Log.DebugFormat("GetAppointments: Booked Resources {0}", tempinfo);
                    }

                    
                    

                    Log.DebugFormat("GetAppointments: Required Attendees  {0}", appointmentDetailed.RequiredAttendees.Count); //v1.7
                    RequiredAttendees = appointmentDetailed.RequiredAttendees.Count;
                    tempinfo = "";
                    foreach (Attendee at in appointmentDetailed.RequiredAttendees)//v1.7
                    {
                        tempinfo = at.Address + "," + tempinfo;
                    }
                    Log.DebugFormat("GetAppointments: Required Attendee {0}", tempinfo);

                    Log.DebugFormat("GetAppointments: Optional Attendees:  {0}", appointmentDetailed.OptionalAttendees.Count); //v1.7
                    OptionalAttendees = appointmentDetailed.OptionalAttendees.Count;
                    tempinfo = "";
                    foreach (Attendee at in appointmentDetailed.OptionalAttendees)//v1.7
                    {
                        tempinfo = at.Address + "," + tempinfo;
                    }
                    Log.DebugFormat("GetAppointments: Optional Attendee {0}", tempinfo);

                    TotalAttendees = OptionalAttendees + RequiredAttendees;

                    string CurrentUrl = appointmentDetailed.JoinOnlineMeetingUrl;
                    string extractedUrl = CurrentUrl.Replace(DefaultDialHttp,"");
                    string[] TempURL = extractedUrl.Split('/');
                    conferenceUrl = "conf:sip:" + TempURL[0] + DefaultDomain + ";gruu;opaque=app:conf:focus:id:" + TempURL[1] + "?"; // in the form of "conf:sip:wewa@microsoft.com;gruu;opaque=app:conf:focus:id:4FNRHN16";
                    MeetingSubject = a.Subject.ToString();
                    
                    if (StartTime < 0)
                    {
                        double temptime = a.Start.Subtract(CurrentTime).TotalMinutes;
                        
                        if (temptime < 30)
                        {
                            Log.DebugFormat("GetAppointments: Starting Early {0}", temptime);
                            MeetingSleep = a.End.Subtract(CurrentTime).TotalMinutes + 1; //v1.5
                            JoinMeeting(conferenceUrl);
                            break;
                        }
                    }
                    else if (StartTime == 0)
                    {
                        Log.DebugFormat("GetAppointments: Right Now");
                        MeetingSleep = a.End.Subtract(CurrentTime).TotalMinutes + 1; //v1.5;
                        JoinMeeting(conferenceUrl);
                        break;
                        
                    }
                    else if (StartTime > 0 && FinishTime > 0)
                    {
                        Log.DebugFormat("GetAppointments: In progress");
                        MeetingSleep = a.End.Subtract(CurrentTime).TotalMinutes + 1; //v1.5;
                        JoinMeeting(conferenceUrl);
                        break;
                        
                    }
                    else if (FinishTime == 0 || FinishTime < 0)
                        Log.DebugFormat("GetAppointments: Meeting Finished");
                        
                }
                
                

            }
        }

        private void JoinMeeting(string conferenceUrl)
        {
            //Process.Start(appointmentDetailed.JoinOnlineMeetingUrl);
            
            Log.DebugFormat("JoinMeeting: Starting {0}" ,conferenceUrl);
            

            if (CurrentEmailAddress.EmailAddress.ToString().Substring(0,3).ToUpper().Contains("LRS"))
            {
                //Process.Start("SfB Meeting Launcher.exe");
                Log.DebugFormat("JoinMeeting: SfB Meeting Launcher starting");
                SfBMeetingLauncher.Start(); //v1.4
            }
            GetAppointmentTimer.Stop();
            try { GetAppointmentTimer.Interval = Convert.ToInt32(MeetingSleep) * 60000; }
            catch { GetAppointmentTimer.Interval = 1 * 60000; } //15 Minutes 
            Log.DebugFormat("JoinMeeting: Timer Interval '{0}' minutes", GetAppointmentTimer.Interval / 60000);

            GetAppointmentTimer.Start();

            try
            {
                _conversation = lyncClient.ConversationManager.JoinConference(conferenceUrl);
                Log.DebugFormat("JoinMeeting: Conversation Attempt 1 - {0}", _conversation.State);

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "Join meeting failed", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                //MessageBox.Show(ex.Message, "Join meeting failed");
                Log.ErrorFormat("JoinMeeting: Failed {0} ", ex.Message);
            }
            Thread.Sleep(5000);

            //Check for Presenter Level status v1.4
            //https://msdn.microsoft.com/en-us/library/office/jj937265.aspx
            //MessageBox.Show(_conversation.SelfParticipant.Properties[ParticipantProperty.IsPresenter].ToString());
            
            if(_conversation.State.ToString() == "Inactive")
            {
                Thread.Sleep(5000);
                _conversation = lyncClient.ConversationManager.JoinConference(conferenceUrl);
                Log.DebugFormat("JoinMeeting: Conversation Attempt 2 - {0}", _conversation.State);
            }
            else
            {
                Log.DebugFormat("JoinMeeting: Conversation State - {0}", _conversation.State);
            }

            if ((bool)_conversation.SelfParticipant.Properties[ParticipantProperty.IsPresenter])
            {
                Log.DebugFormat("JoinMeeting: Presenter Check is true.");
                //ShareWhiteBoard();
                SharePPT();
            }
            else
            {
                Log.DebugFormat("JoinMeeting: Presenter Check is false.", (bool)_conversation.SelfParticipant.Properties[ParticipantProperty.IsPresenter]);
                //_conversation.SelfParticipant.BeginSetProperty(ParticipantProperty.IsPresenter, true, (ar) => { _conversation.SelfParticipant.EndSetProperty(ar); }, null);
                try { SfBMeetingLauncher.Kill(); }
                catch { };
                
            }
            //Start Video v1.4
            try
            {
                var avModality = (AVModality)_conversation.Modalities[ModalityTypes.AudioVideo];
                Log.DebugFormat("JoinMeeting: Start Video {0}", avModality.VideoChannel.State.ToString());
                if(avModality.VideoChannel.State.ToString() == "SendReceive")
                {
                    //nothing
                }
                else
                {
                    StartOurVideo(avModality);
                }
                
            }
            catch { };
            try { SfBMeetingLauncher.Kill(); }
            catch { }
            
            
        }

        private void SharePPT()
        {
            string AppSharingTest = null;
            string ContentSharingTest = null;

            //v1.6 Check for File existing
            if (!File.Exists(_PowerPointDeckName))
            {
                Log.ErrorFormat("SharePPT: {0} was not found", _PowerPointDeckName);
                return;
            }
            //v1.7 Check if Rooms/Resources are more than attendees
            //https://stackoverflow.com/questions/23692127/auto-close-message-box
            if (TotalAttendees <= TotalResources)
            {
                Log.Debug("SharePPT: Total Attendees less than TotalResources Prompted");
                DialogResult dialogResult = SharePPTPrompt();
                
                //DialogResult dialogResult = DialogResult.Cancel;
                //MessageBox.Show("hello");

                if (dialogResult == DialogResult.Yes)
                {
                    //Continue
                    Log.Debug("SharePPT: Total Attendees less than TotalResources - Show PowerPoint");
                }
                else if (dialogResult == DialogResult.No)
                {
                    Log.Error("SharePPT: Total Attendees less than TotalResources - No PowerPoint");
                    return;
                }
                else
                {
                    //Continue
                    Log.Debug("SharePPT: Total Attendees less than TotalResources - No Response");
                }

            }

            //int hrReason;
            Log.DebugFormat("SharePPT: Starting");

            try
            {
                AppSharingTest = ((ApplicationSharingModality)_conversation.Modalities[ModalityTypes.ApplicationSharing]).State.ToString();
                Log.DebugFormat("SharePPT: AppShareTest {0}", AppSharingTest);
            }
            catch
            {
                Log.DebugFormat("SharePPT: AppShareTest Error {0}", AppSharingTest);
            }

            try
            {
                ContentSharingTest = ((ContentSharingModality)_conversation.Modalities[ModalityTypes.ContentSharing]).ActiveContent.Title.ToString();
                Log.DebugFormat("SharePPT: ContentSharingTest {0}", ContentSharingTest);
            }
            catch
            {
                Log.DebugFormat("SharePPT: ContentSharingTest Error {0}", ContentSharingTest);
            }

            //ShareableContent shareableContent = ((ContentSharingModality)_conversation.Modalities[ModalityTypes.ContentSharing]).ActiveContent;

            try
            {
                if (AppSharingTest == "Connected" || ContentSharingTest != null)
                {
                    Log.DebugFormat("SharePPT: Someone already sharing -  {0}", AppSharingTest + "&" + ContentSharingTest);
                }
                else
                {
                    Log.DebugFormat("SharePPT: StartShare Attempt 1");
                    AddPowerPoint();
                    Thread.Sleep(5000);
                    StartContentShare();
                }
            }
            catch
            {
                Log.DebugFormat("SharePPT: StartShare Attempt 2");
                AddPowerPoint();
                Thread.Sleep(5000);
                StartContentShare();
            }

        }

        private void AddPowerPoint()
        {
            if (_conversation.Modalities[ModalityTypes.ContentSharing].State == ModalityState.Disconnected)
            {
                Log.DebugFormat("AddPowerPoint: Disconnected");
                ConnectCSM();

            }
            else
            {
                if (!_conversation.Modalities[ModalityTypes.ContentSharing].CanInvoke(ModalityAction.CreateShareablePowerPointContent))
                {
                    Log.DebugFormat("AddPowerPoint: Return");
                    return;
                }
                Log.DebugFormat("AddPowerPoint: Length {0}", _PowerPointDeckName.Length);
                if (_PowerPointDeckName.Length > 0)
                {
                    try
                    {

                        ContentSharingModality c = ((ContentSharingModality)_conversation.Modalities[ModalityTypes.ContentSharing]);

                        string pptTitle = string.Empty;
                        Int32 lastWhack = _PowerPointDeckName.LastIndexOf(@"\");

                        //Strip off PowerPoint deck path, leaving only the file name to 
                        //be assigned as the ShareableContent.Title property
                        pptTitle = _PowerPointDeckName.Substring(lastWhack + 1, _PowerPointDeckName.Length - (lastWhack + 1));

                        c.BeginCreateContentFromFile(
                            ShareableContentType.PowerPoint,
                            pptTitle,
                            _PowerPointDeckName,
                            false,
                            CreateShareableContentCallback,
                            c);
                        Log.DebugFormat("AddPowerPoint: PowerPoint Upload {0}", _PowerPointDeckName + " " + pptTitle);
                    }
                    catch (InvalidStateException)
                    {
                        MessageBox.Show("Invalid state exception on BeginCreateContent");
                        Log.ErrorFormat("AddPowerPoint: Invalid state exception on BeginCreateContent");
                    }
                    catch (NotInitializedException)
                    {
                        MessageBox.Show("Not initialized exception on BeginCreateContent");
                        Log.ErrorFormat("AddPowerPoint: Not initialized exception on BeginCreateContent");
                    }
                    finally
                    {
                        //Clear the powerpoint deck name. The name will be filled again by 
                        //the user when they choose another deck to upload.
                        //_PowerPointDeckName = null;

                    }
                }
            }
        }

        private static void StartOurVideo(AVModality avModality) //v1.4
        {
            var channelStream = avModality.VideoChannel;
            Log.DebugFormat("StartOurVideo: {0}", channelStream.State.ToString());
            //while (!channelStream.CanInvoke(ChannelAction.Start))
            //{
            //}

            channelStream.BeginStart(ar => { }, channelStream);
            var count = 0;
            while ((channelStream.State != ChannelState.SendReceive) && (count < 5))
            {
                Thread.Sleep(1000);

                try
                {
                    channelStream.BeginStart(ar => { }, channelStream);
                }
                catch (NotSupportedException)
                {
                    //This is normal...
                }
                count++;
            }
        }

        private void AppointmentRecheck(object sender, EventArgs e)
        {
            Log.DebugFormat("AppointmentRecheck: Started");
            GetAppointments();
        }

        private void ConnectCSM()
        {
            Log.DebugFormat("ConnectCSM: Start  {0}", _conversation.Modalities[ModalityTypes.ContentSharing]
            .CanInvoke(ModalityAction.Connect).ToString());
            if (_conversation.Modalities[ModalityTypes.ContentSharing]
            .CanInvoke(ModalityAction.Connect))
            {
                try
                {
                    _conversation.Modalities[ModalityTypes.ContentSharing].
                        BeginConnect((ar) =>
                        {
                            _conversation.Modalities[ModalityTypes.ContentSharing].
                                EndConnect(ar);
                        }
                        , null);
                    Log.DebugFormat("ConnectCSM: Try Finished {0}", _conversation.Modalities[ModalityTypes.ContentSharing].State.ToString());
                }
                catch (OperationException oe)
                {
                    Log.ErrorFormat("ConnectCSM: Error {0}", oe.Message.ToString());
                }
            }
        }

        private void ConversationReCheck(object sender, EventArgs e) //v1.5
        {
            Log.DebugFormat("ConversationRecheck: Started");
            lyncClient.ConversationManager.ConversationAdded += new EventHandler<ConversationManagerEventArgs>(ConversationManager_ConversationAdded);
            lyncClient.ConversationManager.ConversationRemoved += new EventHandler<ConversationManagerEventArgs>(ConversationManager_ConversationRemoved);
        }
        
        #region Lync Event Handlers.

        private void CreateShareableContentCallback(System.IAsyncResult ar)
        {
            try
            {
                ContentSharingModality shareModality = (ContentSharingModality)ar.AsyncState;
                ShareableContent sContent;
                if (_PowerPointDeckName == string.Empty && _NativeFileNameAndPath == string.Empty)
                {
                    sContent = shareModality.EndCreateContent(ar);
                }
                else
                {
                    sContent = shareModality.EndCreateContentFromFile(ar);
                    //_PowerPointDeckName = string.Empty;
                    _NativeFileNameAndPath = string.Empty;
                }
                sContent.Upload();
            }
            catch (MaxContentsExceededException)
            {
                Log.Error("CreateShareableContentCallback: Too many items in content bin");
            }
            catch (ContentTitleExistException)
            {
                Log.Error("CreateShareableContentCallback: Duplicate content title");
            }
            catch (ContentTitleInvalidException)
            {
                Log.Error("CreateShareableContentCallback: Invalid character in content title");
            }
            
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            Log.DebugFormat("RedirectionUrlValidationCallback: Started {0}", redirectionUrl);
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);


            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }

            return result;
        }

        void ConversationManager_ConversationAdded(object sender, ConversationManagerEventArgs e)
        {
            Log.DebugFormat("ConversationManager_ConversationAdded: Start 'e' {0}", e.Conversation);
            Log.DebugFormat("ConversationManager_ConversationAdded: Start '_conversation' {0}", _conversation);
            if (_conversation == null) //v1.5
            {
                _conversation = e.Conversation;
            }

            if (IsShadowConversation(e.Conversation))
            {
                Log.DebugFormat("ConversationManager_ConversationAdded: Shadow {0}", e.Conversation);
                return;
            }

            if (e.Conversation != _conversation)
            {
                //v1.5
                var avModality = e.Conversation.Modalities[ModalityTypes.AudioVideo];
                var avModalityState = avModality.State;
                //Log.DebugFormat("Conversation Added Merge, AV Modality: {0}", avModalityState);
                AcceptVideoWhenVideoAdded(e.Conversation); //v1.5
                Thread.Sleep(5000);
                Log.DebugFormat("ConversationManager_ConversationAdded: Merging - Not Current Conversation");
                Log.DebugFormat("ConversationManager_ConversationAdded: Merging - New Meeting  {0}", e.Conversation.CanInvoke(ConversationAction.Merge).ToString()); //false
                Log.DebugFormat("ConversationManager_ConversationAdded: Merging - Existing Meeting  {0}", _conversation.CanInvoke(ConversationAction.Merge).ToString()); //true
                

                if (_conversation.CanInvoke(ConversationAction.Merge))
                {
                    Log.DebugFormat("ConversationManager_ConversationAdded: Attempt Merge");
                    try
                    {
                        _conversation.BeginMerge(e.Conversation, ModalityTypes.AudioVideo, (ar) =>
                        {
                            _conversation.EndMerge(ar);
                        }, null);

                        _conversation.BeginMerge(e.Conversation, ModalityTypes.ApplicationSharing, (ar) =>
                        {
                            _conversation.EndMerge(ar);
                        }, null);

                        _conversation.BeginMerge(e.Conversation, ModalityTypes.ContentSharing, (ar) =>
                        {
                            _conversation.EndMerge(ar);
                        }, null);

                        
                    }
                    catch (LyncClientException lce)
                    {
                        Log.ErrorFormat("ConversationManager_ConversationAdded: Lync Error:  {0}", lce);
                    }
                    catch (Exception ce)
                    {
                        Log.ErrorFormat("ConversationManager_ConversationAdded: Lync Error:  { 0}", ce);
                    }
                    Thread.Sleep(5000);

                    e.Conversation.End();
                }

                return;
            }

            Log.DebugFormat("ConversationManager_ConversationAdded: Brand New Conversation");

            //Register for participant added events on the new conversation
            _conversation.ParticipantAdded += _conversation_ParticipantAdded; //v1.6 
            
            
            //_conversation.ParticipantRemoved += _conversation_ParticipantRemoved;
            //_conversation.StateChanged += new EventHandler<ConversationStateChangedEventArgs>(_conversation_StateChanged);

            _conversation.Modalities[ModalityTypes.ContentSharing].ModalityStateChanged += Modality_ModalityStateChanged;
            _conversation.Modalities[ModalityTypes.ContentSharing].ActionAvailabilityChanged += _sharingModality_ActionAvailabilityChanged;

            ((ContentSharingModality)_conversation.Modalities[ModalityTypes.ContentSharing]).ActiveContentChanged += _sharingModality_ActiveContentChanged;

            ((ContentSharingModality)_conversation.Modalities[ModalityTypes.ContentSharing]).ContentAdded += _sharingModality_ContentAdded;

            ((ContentSharingModality)_conversation.Modalities[ModalityTypes.ContentSharing]).ContentRemoved += _sharingModality_ContentRemoved;

            Log.DebugFormat("ConversationManager_ConversationAdded: Join Conversation");
            try
            {
                SfBMeetingLauncher.Start();
                Thread.Sleep(2000);
                Log.DebugFormat("ConversationManager_ConversationAdded: SfB Launcher started");
            }
            catch (Exception sfbex)
            {
                Log.ErrorFormat("AcceptVideoWhenVideoAdded: SfB Launcher failed to start '{0}'", sfbex);
            }

            //v1.5 Auto Answer
            try
            {
                var avModality = _conversation.Modalities[ModalityTypes.AudioVideo];

                //Save the video state so that we avoid wacky things when it changes
                var avModalityState = avModality.State;
                Log.DebugFormat("ConversationManager_ConversationAdded: Add AudioVideo ModalityState {0}", avModalityState);

                AcceptVideoWhenVideoAdded(e.Conversation);

            }
            catch (Exception ex)
            {
                Log.Error("ConversationManager_ConversationAdded: Error in ConversationAdded", ex);
            }

            //Add Conference Info to Chat
            //ChatText(); //v1.6
        }

        private void _MessageModality_InstantMessageReceived(object sender, MessageSentEventArgs e) //v1.6 Disabled
        {
            Log.DebugFormat("_MessageModality_InstantMessageReceived" + e.Text.ToString());
            
            if (e.Text.ToUpper().Contains("!HELP") == true)
            {
                try { ChatText(); }
                catch { };

                ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).
                BeginSendMessage(
                "How to Document HyperLink or text here" + System.Environment.NewLine,
                (ar) =>
                {
                    try
                    {
                        ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).EndSendMessage(ar);
                    }
                    catch (LyncClientException) { Log.Debug("_MessageModality_InstantMessageReceived: How To Message could not be delivered"); }
                }
                , null);
                ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).
                BeginSendMessage(
                "How to Powerpoint HyperLynk or text here" + System.Environment.NewLine,
                (ar) =>
                {
                    try
                    {
                        ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).EndSendMessage(ar);
                    }
                    catch (LyncClientException) { Log.Debug("_MessageModality_InstantMessageReceived: PowerPoint Link message could not be delivered"); }
                }
                , null);
            }
        }

        private void _conversation_ParticipantAdded(object sender, ParticipantCollectionChangedEventArgs e) //v1.6
        {

            int CountParticipant = 0;
            Log.DebugFormat("_conversation_ParticipantAdded: Adding Participant {0}", e.Participant.Properties[ParticipantProperty.Name].ToString());
            Thread.Sleep(2000);
            ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).
            BeginSendMessage(
            e.Participant.Properties[ParticipantProperty.Name].ToString() + " has joined the meeting.",
            (ar) =>
            {
                try
                {
                    ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).EndSendMessage(ar);
                }
                catch (LyncClientException) { Log.Debug("_conversation_ParticipantAdded: New Participant message could not be delivered"); }
            }
            , null);

            foreach (Participant FoundParticipant in _conversation.Participants)
            {
                string nameFound = FoundParticipant.Properties[ParticipantProperty.Name].ToString();
                //Log.DebugFormat("_conversation_ParticipantAdded: Current Participants {0}", nameFound);
                if (nameFound.ToUpper().Contains("ROOM") == true)
                {
                    Log.DebugFormat("_conversation_ParticipantAdded: Learning Room {0}",nameFound.ToString());
                }
                else
                {
                    CountParticipant++;

                }
                
            }

            Log.DebugFormat("_conversation_ParticipantAdded: Participants = {0}", CountParticipant.ToString());
            if (CountParticipant == 2 || CountParticipant % 4 == 0)  //2 People or Every 4 people
            {
                Log.DebugFormat("_conversation_ParticipantAdded: Adding Chat message {0}", CountParticipant.ToString());
                ChatText();

            }
            //if (e.Participant.IsSelf == false)
            //{
            //    if (((Microsoft.Lync.Model.Conversation.Conversation)sender).Modalities.ContainsKey(ModalityTypes.InstantMessage))
            //    {
            //        Log.DebugFormat("_conversation_ParticipantAdded: User {0}", e.Participant.Properties[ParticipantProperty.Name].ToString());
                    //((InstantMessageModality)e.Participant.Modalities[ModalityTypes.InstantMessage]).InstantMessageReceived += new EventHandler<MessageSentEventArgs>(_MessageModality_InstantMessageReceived);
                    
             //   }
            //}
        }

        void ConversationManager_ConversationRemoved(object sender, ConversationManagerEventArgs e)
        {
            Log.DebugFormat("ConversationManager_ConversationRemoved: Started");
            if (_conversation != null)
            {
                if (e.Conversation.Equals(_conversation))
                {
                    _conversation = null;
                }
            }

        }

        void Modality_ModalityStateChanged(object sender, ModalityStateChangedEventArgs e)
        {
            Log.DebugFormat("Modality_ModalityStateChanged: Started {0}", sender.GetType().Name );
            //Modality will be connected for each particpant whethere they have accepted 
            //the shariing invite or not.

            switch (sender.GetType().Name)
            {
                case "ContentSharingModality":
                    ContentSharingModality thisModality = sender as ContentSharingModality;
                    if (thisModality == ((ContentSharingModality)_conversation.
                        Modalities[ModalityTypes.ContentSharing]))
                    {
                        switch (e.NewState)
                        {
                            case ModalityState.Connecting:
                                //this.Invoke(new ChangeControlBackgroundDelegate(ChangeControlBackground),new object[] { ModalityConnectedState_Panel,System.Drawing.Color.Gold });
                                break;
                            case ModalityState.Connected:
                                Log.DebugFormat("Modality_ModalityStateChanged: Connected");
                                break;
                            case ModalityState.Disconnected:
                                break;
                        }
                    }
                    break;
            }
        }

        void _sharingModality_ActionAvailabilityChanged(object sender, ModalityActionAvailabilityChangedEventArgs e)
        {
            Log.DebugFormat("Modality_ModalityStateChanged: Started");

            ContentSharingModality csm = ((ContentSharingModality)_conversation.Modalities[ModalityTypes.ContentSharing]);

            switch (e.Action)
            {
                case ModalityAction.CreateShareableNativeFileOnlyContent:
                    if (e.IsAvailable == false)
                        return;

                    if (_NativeFileName.Length == 0 || _NativeFileNameAndPath.Length == 0)
                    {
                        return;
                    }
                    UploadAFile();
                    break;

                case ModalityAction.CreateShareablePowerPointContent:
                    //If the user has chosen a PowerPoint deck to upload, then create 
                    //and upload it when the action is available
                    if (_PowerPointDeckName.Length > 0)
                    {
                        try
                        {
                            string pptTitle = string.Empty;
                            Int32 lastWhack = _PowerPointDeckName.LastIndexOf(@"\");

                            //Strip off PowerPoint deck path, leaving only the file name to 
                            //be assigned as the ShareableContent.Title property
                            pptTitle = _PowerPointDeckName.Substring(lastWhack + 1, _PowerPointDeckName.Length - (lastWhack + 1));
                            Log.DebugFormat("_sharingModality_ActionAvailabilityChanged : {0}", pptTitle);
                            csm.BeginCreateContentFromFile(
                                ShareableContentType.PowerPoint,
                                pptTitle,
                                _PowerPointDeckName,
                                false,
                                (ar) =>
                                {
                                    try
                                    {
                                        ShareableContent sContent;
                                        sContent = csm.EndCreateContentFromFile(ar);
                                        //_PowerPointDeckName = string.Empty;
                                        sContent.Upload();

                                    }
                                    catch (MaxContentsExceededException)
                                    {
                                        Log.ErrorFormat("The meeting content bin is full. PowerPoint deck {0} cannot be uploaded", _PowerPointDeckName);
                                    }
                                    catch (ContentTitleExistException)
                                    {
                                        Log.ErrorFormat("The meeting content bin already contains an item with this title. PowerPoint deck {0} cannot be uploaded", _PowerPointDeckName);
                                    }
                                    catch (ContentTitleInvalidException)
                                    {
                                        Log.ErrorFormat("The whiteboard title is invalid. PowerPoint deck {0} cannot be uploaded", _PowerPointDeckName);
                                    }
                                }
                                , null);

                        }
                        catch (InvalidStateException) { }
                        catch (NotInitializedException) { }
                        finally
                        {
                            //Clear the powerpoint deck name. The name will be filled again 
                            //by the user when they choose another deck to upload.
                            //_PowerPointDeckName = null;
                        }
                    }


                    break;

                //case ModalityAction.CreateShareableWhiteboardContent:
                    //if (csm.CanInvoke(ModalityAction.CreateShareableWhiteboardContent))
                    //{
                    //    csm.BeginCreateContent(ShareableContentType.Whiteboard, "WhiteBoard",(ar) =>
                    //        {
                    //            try
                    //            {
                    //                ShareableContent sContent = csm.EndCreateContent(ar);
                    //                sContent.Upload();
                    //            }
                    //            catch (MaxContentsExceededException)
                    //            {
                    //                MessageBox.Show(
                    //                    "The meeting content bin is full. Whiteboard cannot be "
                    //                    + "uploaded");
                    //            }
                    //            catch (ContentTitleExistException)
                    //            {
                    //                MessageBox.Show(
                    //                    "The meeting content bin already contains an item with "
                    //                    + "this title. Whiteboard cannot be uploaded");
                    //            }
                    //            catch (ContentTitleInvalidException)
                    //            {
                    //                MessageBox.Show(
                    //                    "The whiteboard title is invalid. Whiteboard cannot "
                    //                    + "be uploaded");
                    //            }
                    //        }
                    //    , null);

                        //Clear the whiteboard title entry field
                        //this.Invoke(new ChangeControlTextDelegate(ChangeControlText), new object[] { ShareableContentTitle_Textbox, "" });
                    //}

                    //break;

                case ModalityAction.Accept:
                    csmSharing = (ContentSharingModality)sender;
                    //this.Invoke(d, new object[] { Accept_Button, e.IsAvailable });
                    break;

                case ModalityAction.Reject:
                    //this.Invoke(d, new object[] { Reject_Button, e.IsAvailable });
                    break;
            }
        }

        void _sharingModality_ContentAdded(object sender, ContentCollectionChangedEventArgs e)
        {
            Log.DebugFormat("_sharingModality_ContentAdded: Started");

        }

        void _sharingModality_ActiveContentChanged(object sender, ActiveContentChangedEventArgs e)
        {
            Log.DebugFormat("_sharingModality_ActiveContentChanged: Started");
            if (e.ActiveContent != null)
            {
                _ShareableContent = e.ActiveContent;

                if (e.ActiveContent.Type == ShareableContentType.PowerPoint)
                {

                    Int32 hrReason;
                    if (e.ActiveContent.CanInvoke(ShareableContentAction.ClearAllAnnotations, out hrReason))
                    {

                    }

                }
                else if (e.ActiveContent.Type == ShareableContentType.Whiteboard)
                {
                    Int32 hrReason;
                    if (e.ActiveContent.CanInvoke(ShareableContentAction.ClearAllAnnotations, out hrReason))
                    {

                    }
                    if (e.ActiveContent.CanInvoke(ShareableContentAction.SaveAnnotation, out hrReason))
                    {

                    }

                }
            }
        }

        void _sharingModality_ContentRemoved(object sender, ContentCollectionChangedEventArgs e)
        {
            Log.DebugFormat("_sharingModality_ContentRemoved: Started");
            
            foreach (ShareableContent sc in ((ContentSharingModality)_conversation.Modalities[ModalityTypes.ContentSharing]).ContentCollection)
            {
                //string newItem = sc.Type.ToString() + "-" + sc.Title;
                //this.Invoke(new AddAListItemDelegate(AddAListItem), new object[] { Content_Listbox, newItem });
            }
        }

        private Boolean IsShadowConversation(Microsoft.Lync.Model.Conversation.Conversation newConversation)
        {
            Log.DebugFormat("IsShadowConversation: Started");
            if (newConversation.SelfParticipant != null)
            {
                Log.DebugFormat("IsShadowConversation: Return False");
                return false;
            }
            foreach (Modality m in newConversation.Modalities.Values)
            {
                if (m.State == ModalityState.Notified)
                {
                    Log.DebugFormat("IsShadowConversation: ModalityState.Notified");
                    return false;
                }
            }
            Log.DebugFormat("IsShadowConversation: Return True");
            return true;
        }

        #endregion

        private void AcceptVideoWhenVideoAdded(Microsoft.Lync.Model.Conversation.Conversation conversation)
        {
            var avModality = (AVModality)conversation.Modalities[ModalityTypes.AudioVideo];

            //Check if the new conversation is a new incoming video request
            if (avModality.State == ModalityState.Notified && AutoAnswer())
                AnswerVideo(conversation);

            avModality.ModalityStateChanged += (o, args) =>
            {
                try
                {
                    var newState = args.NewState;
                    Log.DebugFormat("AcceptVideoWhenVideoAdded: Conversation Modality State Changed to '{0}'", newState);

                    if (newState == ModalityState.Notified && AutoAnswer())
                        AnswerVideo(conversation);
                    if (newState == ModalityState.Connected && AutoAnswer())
                        StartOurVideo(avModality);
                    if (newState == ModalityState.Disconnected && AutoAnswer()) //v1.5
                    {
                        if (!_conversation.CanInvoke(ConversationAction.Merge))
                        {
                            //Log.DebugFormat("AcceptVideoWhenVideoAdded: Ending Current Conversation");
                            //v1.6 conversation.End();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("AcceptVideoWhenVideoAdded: Error handling modality state change ", ex);
                }
            };
        }

        private static void AnswerVideo(Microsoft.Lync.Model.Conversation.Conversation conversation) //v1.5
        {
            Log.DebugFormat("AnswerVideo: Started");
            var converstationState = conversation.State;
            if (converstationState == ConversationState.Terminated)
            {
                return;
            }

            var av = (AVModality)conversation.Modalities[ModalityTypes.AudioVideo];
            if (av.CanInvoke(ModalityAction.Connect))
            {
                av.Accept();

                // Get ready to be connected, then WE can start OUR video
                //av.ModalityStateChanged += AVModality_ModalityStateChanged;
            }
            else
            {
                Log.Warn("AnswerVideo: Unable to start video do to 'CanInvoke' being false");
            }
        }

        private void ChatText() //v1.6
        {
            //v1.6 https://msdn.microsoft.com/en-us/library/office/jj937289.aspx
            Thread.Sleep(5000);
            ConferenceAccessInformation conferenceAccess = (ConferenceAccessInformation)_conversation.Properties[ConversationProperty.ConferenceAccessInformation]; //v1.6

            Log.DebugFormat("ChatText: conferenceAccess {0}", conferenceAccess.ToString());
            StringBuilder MeetingKey = new StringBuilder();
            try
            {
                //These properties are used to invite people by creating an email (or text message, or IM)
                //and adding the dial in number, external Url, internal Url, and conference Id

                if (conferenceAccess.Id.Length > 0)
                {
                    Log.DebugFormat("ChatText: conferenceAccess {0}", conferenceAccess.Id);
                    MeetingKey.Append("Meeting Id: " + conferenceAccess.Id);
                    MeetingKey.Append(System.Environment.NewLine);
                }

                if (conferenceAccess.AdmissionKey.Length > 0)
                {
                    Log.DebugFormat("ChatText: conferenceAccess {0}", conferenceAccess.AdmissionKey);
                    MeetingKey.Append(conferenceAccess.AdmissionKey);
                    MeetingKey.Append(System.Environment.NewLine);
                }

                string[] attendantNumbers = (string[])conferenceAccess.AutoAttendantNumbers;

                StringBuilder sb2 = new StringBuilder();
                sb2.Append(System.Environment.NewLine);
                foreach (string aNumber in attendantNumbers)
                {
                    //Console.WriteLine("Conference Numbers: " + aNumber);
                    sb2.Append(aNumber);
                    sb2.Append(System.Environment.NewLine);
                }
                if (sb2.ToString().Trim().Length > 0)
                {
                    MeetingKey.Append("Dial in numbers:" + sb2.ToString());
                    MeetingKey.Append(System.Environment.NewLine);
                }

                // if (conferenceAccess.ExternalUrl.Length > 0)
                //{
                //    MeetingKey.Append("External Url: " + conferenceAccess.ExternalUrl);
                //    MeetingKey.Append(System.Environment.NewLine);
                //}

                //if (conferenceAccess.InternalUrl.Length > 0)
                //{
                //    MeetingKey.Append("Inner Url: " + conferenceAccess.InternalUrl);
                //    MeetingKey.Append(System.Environment.NewLine);
                //}
                Log.DebugFormat("ChatText: MeetingKey {0}", MeetingKey.ToString());


            }
            catch (System.NullReferenceException nr)
            {
                Log.ErrorFormat("ChatText: NullReference {0}", nr.Message);

            }
            catch (LyncClientException lce)
            {
                Log.ErrorFormat("ChatText: ConferenceAccessInformation changed {0}", lce.Message);

            }
            catch (Exception me)
            {
                Log.ErrorFormat("ChatText: Meeting Issue {0}", me.Message);

            }

            if (MeetingKey != null)
            {
                ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).
            BeginSendMessage(
            "Hello. If you are having audio issues, you can call this number as a last resort." + System.Environment.NewLine + MeetingKey.ToString(),
            (ar) =>
            {
                try
                {
                    ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).EndSendMessage(ar);
                }
                catch (LyncClientException) { Log.Debug("ChatText: Message could not be delivered"); }
            }
            , null);
            }

    ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).
        BeginSendMessage(
        "How to Document - Create Link Here" + System.Environment.NewLine,
        (ar) =>
        {
            try
            {
                ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).EndSendMessage(ar);
            }
            catch (LyncClientException) { Log.Debug("_MessageModality_InstantMessageReceived: How To Message could not be delivered"); }
        }
        , null);
            ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).
            BeginSendMessage(
            "How to Powerpoint - Create Link Here" + System.Environment.NewLine,
            (ar) =>
            {
                try
                {
                    ((InstantMessageModality)_conversation.Modalities[ModalityTypes.InstantMessage]).EndSendMessage(ar);
                }
                catch (LyncClientException) { Log.Debug("_MessageModality_InstantMessageReceived: PowerPoint Link message could not be delivered"); }
            }
            , null);

        }

        private void UploadAFile()
        {
            Log.DebugFormat("UploadAFile: Started");
            try
            {
                if (((ContentSharingModality)_conversation.
                    Modalities[ModalityTypes.ContentSharing]).
                    CanInvoke(ModalityAction.CreateShareableNativeFileOnlyContent))
                {
                    ContentSharingModality contentSharingModality =
                        (ContentSharingModality)_conversation.
                        Modalities[ModalityTypes.ContentSharing];

                    contentSharingModality.BeginCreateContentFromFile(
                        ShareableContentType.NativeFile,
                        _NativeFileName,
                        _NativeFileNameAndPath,
                        true,
                        (ar) =>
                        {
                            ShareableContent sContent = contentSharingModality.
                                EndCreateContentFromFile(ar);
                            //_PowerPointDeckName = string.Empty;
                            _NativeFileNameAndPath = string.Empty;
                            sContent.Upload();
                        }
                        , null);
                }
            }
            catch (InvalidStateException)
            {
                MessageBox.Show("Invalid state exception on BeginCreateContent ");
                Log.Error("UploadAFile: Invalid state exception on BeginCreateContent");
            }
            catch (NotInitializedException)
            {
                MessageBox.Show("Not initialized exception on BeginCreateContent ");
                Log.Error("UploadAFile: Not initialized exception on BeginCreateContent");
            }
            finally { _NativeFileNameAndPath = string.Empty; }

        }

        private DialogResult SharePPTPrompt() //v1.7
        {
            // Wait for some result or make the default decision
            //https://www.nuget.org/packages/AutoClosingMessageBox/1.0.0.2
            var result = AutoClosingMessageBox.Show(
                        text: "Do you want to present the PowerPoint help file?",
                        caption: "Share Help PowerPoint",
                        timeout: 10000,
                        buttons: MessageBoxButtons.YesNoCancel,
                        defaultResult: DialogResult.Cancel);
            Log.DebugFormat("SharePPTPrompt: {0}", result.ToString());
            return result;

        }
    }
}

