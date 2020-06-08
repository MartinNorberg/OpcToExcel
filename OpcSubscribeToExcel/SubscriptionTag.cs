namespace OpcSubscribeToExcel
{
    using Opc;
    using Opc.Da;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using System.Text;
    using System.Threading.Tasks;

    public class SubscriptionTag : ItemIdentifier, INotifyPropertyChanged
    {
        private System.Type reqType;
        private int maxAge;
        private bool maxAgeSpecified;
        private bool active;
        private bool activeSpecified;
        private float deadband;
        private bool deadbandSpecified;
        private int samplingRate;
        private bool samplingRateSpecified;
        private bool enableBuffering;
        private bool enableBufferingSpecified;
        private string value;
        private string timestamp;
        public SubscriptionTag()
        {

        }
        

        public SubscriptionTag(ItemIdentifier item)
        {

        }

        public SubscriptionTag(Item item)
        {

        }

        public System.Type ReqType
        {
            get => this.reqType;
            set
            {
                if (ReferenceEquals(value, this.reqType))
                {
                    return;
                }

                this.reqType = value;
                this.OnPropertyChanged();
            }
        }
        public int MaxAge
        {
            get => this.maxAge;
            set
            {
                if (value == this.maxAge)
                {
                    return;
                }

                this.maxAge = value;
                this.OnPropertyChanged();
            }
        }
        public bool MaxAgeSpecified
        {
            get => this.maxAgeSpecified;
            set
            {
                if (value == this.maxAgeSpecified)
                {
                    return;
                }

                this.maxAgeSpecified = value;
                this.OnPropertyChanged();
            }
        }
        public bool Active
        {
            get => this.active;
            set
            {
                if (value == this.active)
                {
                    return;
                }

                this.active = value;
                this.OnPropertyChanged();
            }
        }
        public bool ActiveSpecified
        {
            get => this.activeSpecified;
            set
            {
                if (value == this.activeSpecified)
                {
                    return;
                }

                this.activeSpecified = value;
                this.OnPropertyChanged();
            }
        }
        public float Deadband
        {
            get => this.deadband;
            set
            {
                if (value == this.deadband)
                {
                    return;
                }

                this.deadband = value;
                this.OnPropertyChanged();
            }
        }
        public bool DeadbandSpecified
        {
            get => this.deadbandSpecified;
            set
            {
                if (value == this.deadbandSpecified)
                {
                    return;
                }

                this.deadbandSpecified = value;
                this.OnPropertyChanged();
            }
        }
        public int SamplingRate
        {
            get => this.samplingRate;
            set
            {
                if (value == this.samplingRate)
                {
                    return;
                }

                this.samplingRate = value;
                this.OnPropertyChanged();
            }
        }
        public bool SamplingRateSpecified
        {
            get => this.samplingRateSpecified;
            set
            {
                if (value == this.samplingRateSpecified)
                {
                    return;
                }

                this.samplingRateSpecified = value;
                this.OnPropertyChanged();
            }
        }
        public bool EnableBuffering
        {
            get => this.enableBuffering;
            set
            {
                if (value == this.enableBuffering)
                {
                    return;
                }

                this.enableBuffering = value;
                this.OnPropertyChanged();
            }
        }
        public bool EnableBufferingSpecified
        {
            get => this.enableBufferingSpecified;
            set
            {
                if (value == this.enableBufferingSpecified)
                {
                    return;
                }

                this.enableBufferingSpecified = value;
                this.OnPropertyChanged();
            }
        }
        public string Timestamp
        {
            get => this.timestamp;
            set
            {
                if (value == this.timestamp)
                {
                    return;
                }

                this.timestamp = value;
                this.OnPropertyChanged();
            }
        }
        public string Value
        {
            get => this.value;
            set
            {
                if (value == this.value)
                {
                    return;
                }

                this.value = value;
                this.OnPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null) => this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
