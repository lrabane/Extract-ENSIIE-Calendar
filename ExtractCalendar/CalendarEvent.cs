using System;
using System.Collections.Generic;
using System.Text;

namespace ExtractCalendar
{
    class CalendarEvent
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Start { get; set; }
        public string End { get; set; }
        public bool AllDay { get; set; }
        public bool Editable { get; set; }
    }
}
