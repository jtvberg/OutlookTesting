namespace OutlookTesting
{
    using Microsoft.Office.Interop.Outlook;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Data;
    using System.Windows.Input;
    using System.Windows.Media;
    using System.Windows.Threading;

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.SizeToContent = System.Windows.SizeToContent.WidthAndHeight;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            this.GetAllCalendarItems();
            DispatcherTimer timer = new DispatcherTimer(new TimeSpan(0, 0, 1), DispatcherPriority.Normal, delegate { this.BasicChecks(this.cil); }, this.Dispatcher);
        }

        private List<ICalendarItems> cil;

        public void GetAllCalendarItems()
        {
            var bd = DateTime.Today.ToShortDateString();
            var ed = DateTime.Today.Date.AddDays(7).ToShortDateString();
            var a = new Microsoft.Office.Interop.Outlook.Application();
            Items i = a.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Items;
            i.IncludeRecurrences = true;
            i.Sort("[Start]");
            i = i.Restrict("[Start] >= '" + bd + "' AND [Start] < '" + ed + "'");
            var r =
                from ai in i.Cast<AppointmentItem>()
                select new CalendarItem() 
                { Start = ai.Start,
                    End = ai.End,
                    Subject = ai.Subject,
                    AllDayEvent = ai.AllDayEvent,
                    Body = ai.Body,
                    Duration = ai.Duration,
                    Organizer = ai.Organizer,
                    Recurring = ai.IsRecurring,
                    Location = ai.Location,
                    Type = ai.MeetingStatus.ToString(),
                    Recps = ai.Recipients,
                    RecipientCount = ai.Recipients.Count };
            this.cil = new List<ICalendarItems>(r.ToList<ICalendarItems>());
            this.BasicChecks(this.cil);
            this.NextCheck(this.cil);
            this.AddDays(this.cil);
            this.LvOutlookItems.ItemsSource = this.cil;
        }

        private void AddItem(Microsoft.Office.Interop.Outlook.Application a)
        {
            AppointmentItem newAppointment = (AppointmentItem)
            a.CreateItem(OlItemType.olAppointmentItem);
            newAppointment.Start = DateTime.Now.AddHours(2);
            newAppointment.End = DateTime.Now.AddHours(3);
            newAppointment.Location = "";
            newAppointment.Body = "";
            newAppointment.AllDayEvent = false;
            newAppointment.Subject = "Unavailable";
            newAppointment.Save();
            newAppointment.Display(true);
        }

        private void AddDays(List<ICalendarItems> cis)
        {
            var d = new DateTime();
            for (int i = 0; i < cis.Count; i++)
            {
                var ci = cis[i];
                if (ci.GetType() == typeof(CalendarItem) && ((CalendarItem)ci).Start.Date > d)
                {
                    this.cil.Insert(i, new CalendarDate() { Day = ((CalendarItem)ci).Start.ToShortDateString() });
                    d = ((CalendarItem)ci).Start.Date;
                }
            }
        }

        private void BasicChecks(List<ICalendarItems> cis)
        {
            foreach (var ct in cis)
            {
                if (ct.GetType() == typeof(CalendarItem))
                {
                    var ci = (CalendarItem)ct;
                    if (ci.Duration >= 240)
                    {
                        ci.LongEvent = true;
                    }
                    else
                    {
                        ci.LongEvent = false;
                    }

                    if (ci.Start < DateTime.Now && ci.End > DateTime.Now)
                    {
                        ci.Current = true;
                    }
                    else
                    {
                        ci.Current = false;
                    }

                    ci.Conflict = false;
                    foreach (var cit in cis)
                    {
                        if (cit.GetType() == typeof(CalendarItem))
                        {
                            var cic = (CalendarItem)cit;
                            if (!ci.LongEvent && !cic.LongEvent && !ci.AllDayEvent && !cic.AllDayEvent && ci.Subject != cic.Subject && ((ci.Start >= cic.Start && ci.Start < cic.End) || (ci.End > cic.Start && ci.End <= cic.End)))
                            {
                                ci.Conflict = true;
                            }
                        }
                    }
                }
            }
        }

        private void NextCheck(List<ICalendarItems> cis)
        {
            foreach (CalendarItem ci in cis)
            {
                ci.Next = false;
                if (ci.Start > DateTime.Now)
                {
                    ci.Next = true;
                    return;
                }
            }
        }

        private void SvMainPreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            ((ScrollViewer)sender).ScrollToVerticalOffset(((ScrollViewer)sender).VerticalOffset + (-e.Delta / 2));
        }

        private void SvMainMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            this.DragMove();
        }

        private void CalMainSelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            foreach (DateTime day in this.CalMain.SelectedDates)
            {
                this.SvMain.ScrollToHome();
                return;
            }
        }
    }

    public interface ICalendarItems { }

    public class CalendarItem : ICalendarItems
    {
        public string Subject { get; set; }

        public DateTime Start { get; set; }

        public DateTime End { get; set; }

        public bool AllDayEvent { get; set; }

        public string Body { get; set; }

        public string Organizer { get; set; }

        public int Duration { get; set; }

        public bool Current { get; set; }

        public bool Next { get; set; }

        public bool Conflict { get; set; }

        public bool LongEvent { get; set; }

        public bool Recurring { get; set; }

        public string Location { get; set; }

        public int RecipientCount { get; set; }

        public Recipients Recps { get; set; }

        public string Type { get; set; }
    }

    public class CalendarDate : ICalendarItems
    {
        public string Day { get; set; }
    }

    public class StartDateTimeToTime : IMultiValueConverter
    {
        public object Convert(object[] value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if((bool)value[1])
            {
                return "All";
            }
            return ((DateTime)value[0]).ToShortTimeString().Replace("AM", string.Empty).Replace("PM", string.Empty).Replace(":00", string.Empty);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }

    public class EndDateTimeToTime : IMultiValueConverter
    {
        public object Convert(object[] value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if ((bool)value[1])
            {
                return "Day";
            }
            return ((DateTime)value[0]).ToShortTimeString().Replace("AM", string.Empty).Replace("PM", string.Empty).Replace(":00", string.Empty);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }

    public class OrganizerToDarkColor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if(value.ToString().Contains("Vandenberg"))
            {
                return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF0a7e07"));
            }
            if (value.ToString().Contains("Durham") || value.ToString().Contains("Vetsch"))
            {
                return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFEF8D31"));
            }

            return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF303f9f"));
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }

    public class OrganizerToLightColor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if (value.ToString().Contains("Vandenberg"))
            {
                return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF5677fc"));
            }
            if (value.ToString().Contains("Durham") || value.ToString().Contains("Vetsch"))
            {
                return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFd84315"));
            }

            return new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF3f51b5"));
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }

    public class ConflictToColor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if ((bool)value)
            {
                return new SolidColorBrush(Colors.Red);
            }
            return new SolidColorBrush(Colors.Transparent);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }

    public class RecurringToColor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if ((bool)value)
            {
                return new SolidColorBrush(Colors.Black);
            }
            return new SolidColorBrush(Colors.Transparent);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }

    public class AllDayEventToColor : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if ((bool)value)
            {
                return new SolidColorBrush(Colors.Black);
            }
            return new SolidColorBrush(Colors.Transparent);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }

    public class CurentToColor : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            bool current = (bool)values[0];
            bool next = (bool)values[1];
            if (current)
            {
                return new SolidColorBrush(Colors.Green);
            }
            if (next)
            {
                return new SolidColorBrush(Colors.Yellow);
            }
            return new SolidColorBrush(Colors.Transparent);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            return null;
        }
    }
}
