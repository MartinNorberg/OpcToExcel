using Opc;
using Opc.Da;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Ookii.Dialogs.Wpf;
using System.IO;
using System.Windows.Threading;
using System.Configuration;
using System.Windows;

namespace OpcSubscribeToExcel
{
    public class ViewModel : INotifyPropertyChanged
    {
        private Opc.Da.Subscription subscription;
        private Opc.Da.Server daServer;
        private string url;
        private string path;
        private string fileName;
        private string connectionStatus;
        private string subscribeText;
        private string lastSaved;
        private bool saveCyclic;
        private bool newFile;
        private float cyclicTime;
        DispatcherTimer timer;

        public ObservableCollection<SubscriptionTag> @Items { get; } = new ObservableCollection<SubscriptionTag>();

        public ObservableCollection<Exception> Exceptions { get; } = new ObservableCollection<Exception>();

        public ViewModel()
        {            
            this.timer = new DispatcherTimer();
            //var testItem = new SubscriptionTag();
            //testItem.ItemName = "Applications.Application_1.Test.DataToExcel.MyBool";
            //this.Items.Add(testItem);
            this.connectionStatus = "Connect";
            this.subscribeText = "Start subscribe";
            this.ConnectCommand = new RelayCommand(_ => TryCatch(() => this.Connect()));
            this.StartCommand = new RelayCommand(_ => TryCatch(() => {
                if (this.subscription != null)
                {                    
                    this.StopSubscribing();
                }
                else
                {
                    this.StartSubscribing();
                }
            }));
            this.BrowseCommand = new RelayCommand(_ => TryCatch(() => this.BrowseLocation()));
            this.SaveConfigCommand = new RelayCommand(_ => TryCatch(() => this.SaveConfig()));
            this.TryCatch(() => this.LoadConfig());
        }



        public event PropertyChangedEventHandler PropertyChanged;

        public ICommand StartCommand { get; }        

        public ICommand ConnectCommand { get; }

        public ICommand BrowseCommand { get; }

        public ICommand SaveConfigCommand { get; }
        
        public string Url
        {
            get => this.url;
            set
            {
                if (value == this.url)
                {
                    return;
                }

                this.url = value;
                this.OnPropertyChanged();
            }
        }

        public string LastSaved
        {
            get => this.lastSaved;
            set
            {
                if (value == this.lastSaved)
                {
                    return;
                }

                this.lastSaved = value;
                this.OnPropertyChanged();
            }
        }

        public string ConnectionStatus
        {
            get => this.connectionStatus;
            set
            {
                if (value == this.connectionStatus)
                {
                    return;
                }

                this.connectionStatus = value;
                this.OnPropertyChanged();
            }
        }

        public string SubscribeText
        {
            get => this.subscribeText;
            set
            {
                if (value == this.subscribeText)
                {
                    return;
                }

                this.subscribeText = value;
                this.OnPropertyChanged();
            }
        }

        public string Path
        {
            get => this.path;
            set
            {
                if (value == this.path)
                {
                    return;
                }

                this.path = value;
                this.OnPropertyChanged();
            }
        }

        public string FileName
        {
            get => this.fileName;
            set
            {
                if (value == this.fileName)
                {
                    return;
                }

                this.fileName = value;
                this.OnPropertyChanged();
            }
        }

        public bool SaveCyclic 
        { 
            get => saveCyclic;
            set
            {
                if (value == this.saveCyclic)
                {
                    return;
                }

                this.saveCyclic = value;
                this.OnPropertyChanged();
            }

        }

        public bool NewFile
        {
            get => newFile;
            set
            {
                if (value == this.newFile)
                {
                    return;
                }

                this.newFile = value;
                this.OnPropertyChanged();
            }
        }

        public float CyclicTime
        {
            get => cyclicTime;
            set
            {
                if (value == this.cyclicTime)
                {
                    return;
                }

                this.cyclicTime = value;
                this.OnPropertyChanged();
                this.timer.Interval = TimeSpan.FromSeconds(this.cyclicTime);

            }
        }

        private void Connect()
        {

            if (this.daServer != null)
            {
                this.daServer.Dispose();
                this.daServer = null;
                this.connectionStatus = "Connect";
                this.OnPropertyChanged(nameof(this.ConnectionStatus));
                return;
            }

            this.daServer = new Opc.Da.Server(new OpcCom.Factory(), new Opc.URL(this.Url));
            this.daServer.Connect();

            this.connectionStatus = "Disconnect";
            this.OnPropertyChanged(nameof(this.ConnectionStatus));

        }

        private void StartSubscribing()
        {

            SubscriptionState state = new SubscriptionState();
            state.Active = true;
            state.ClientHandle = (object)Guid.NewGuid().ToString();
            state.Deadband = 0.0F;
            state.KeepAlive = 0;
            state.Locale = "";
            state.Name = "OpcToExcel";
            state.UpdateRate = 10;            
            subscription = (Subscription)daServer.CreateSubscription(state);

            Item[] opcItems = new Item[this.Items.Count];
            // Add subscription for all tags in transaction 
            int i = 0;
            foreach (var item in this.Items)
            {
                opcItems[i] = new Item(item);
                opcItems[i].ClientHandle = Guid.NewGuid().ToString();
                opcItems[i].Active = true;
                i++;
            }

            // add items to subscription.
            ItemResult[] ires = subscription.AddItems(opcItems);

            // check if Result is "S_OK"
            foreach (ItemResult item in ires)
            {
                // check if TagConfig is supported by Opc
                if (!item.ResultID.Name.Name.Equals("S_OK"))
                {
                    string str = item.Key;

                    if (!item.ItemName.Equals(string.Empty))
                    {
                        throw new Exception($"Failed to add subscription to tag: {item.ItemName} with result code: {item.ResultID.Name.Name}");
                    }
                }
            }

            // register for data updates.
            this.subscription.DataChanged += new DataChangedEventHandler(this.OnDataChange);

            this.subscribeText = "Stop subscribe";
            this.OnPropertyChanged(nameof(this.SubscribeText));


            timer.Interval = TimeSpan.FromSeconds(this.cyclicTime);
            timer.Tick += timer_Tick;
            timer.Start();
        }

        void timer_Tick(object sender, EventArgs e)
        {
            foreach (var item in Items)
            {
                this.TryCatch(() => this.WriteTagsToExcel($"{item.ItemName} ; {item.Value} ; {item.Timestamp}"));
            }

        }

        public void StopSubscribing()
        {
            timer.Tick -= timer_Tick;
            this.subscription.DataChanged -= new DataChangedEventHandler(this.OnDataChange);            
            this.daServer.CancelSubscription(this.subscription);
            this.subscription.Dispose();
            this.subscribeText = "Start subscribe";
            this.OnPropertyChanged(nameof(this.SubscribeText));
        }

        private void BrowseLocation()
        {
            VistaFolderBrowserDialog dialog = new VistaFolderBrowserDialog();
            dialog.Description = "Please select a folder.";
            dialog.UseDescriptionForTitle = true;

            if ((bool)dialog.ShowDialog())
            {
                this.path = dialog.SelectedPath;
                this.OnPropertyChanged(nameof(this.Path));
            }
        }

        private void LoadConfig()
        {
           this.Url = ConfigurationManager.AppSettings["OpcUrl"].ToString();
           this.Path = ConfigurationManager.AppSettings["Path"].ToString();
           this.FileName = ConfigurationManager.AppSettings["FileName"].ToString();
           this.SaveCyclic = bool.Parse(ConfigurationManager.AppSettings["SaveCyclic"].ToString());
           this.NewFile = bool.Parse(ConfigurationManager.AppSettings["NewFile"].ToString());
           this.CyclicTime = float.Parse(ConfigurationManager.AppSettings["CyclicTime"].ToString());
        }

        private void SaveConfig()
        {
            throw new ApplicationException("Function not implemented yet...sorry");
        }

        private void WriteTagsToExcel(string data)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                this.fileName = "OPCExport";
                this.OnPropertyChanged(nameof(this.FileName));
            }


            var file = this.path + $@"\{this.fileName}";
            
            if (this.newFile)
            {
                file += DateTime.Today.ToShortDateString();
            }

            file += ".csv";
            if (!File.Exists(file))
            {
                File.Create(file);
            }
            File.SetAttributes(file, FileAttributes.Normal);
            using (var stream = File.Open(file, FileMode.Append, FileAccess.Write))
            using (var sw = new StreamWriter(stream))
            {
                sw.WriteLine(data);
            }
            this.LastSaved = DateTime.Now.ToString();
                
                
            
        }

        private void OnDataChange(object subscriptionHandle, object requestHandle, ItemValueResult[] values)
        {
            foreach (var valueItem in values)
            {
                var item = this.Items.FirstOrDefault(i => i.ItemName == valueItem.ItemName);
                if (item != null)
                {
                    item.Value = valueItem.Value.ToString();
                    item.Timestamp = valueItem.Timestamp.ToString();
                    if (!this.saveCyclic)
                    {
                       TryCatch(()=> this.WriteTagsToExcel($"{item.ItemName} ; {item.Value} ; {item.Timestamp}"));
                    }
                }
            }
        }

        public void TryCatch(Action action)
        {
            try
            {
                action();
            }
            catch (Exception e)
            {
                this.Exceptions.Add(e);
            }
        }

        private protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
