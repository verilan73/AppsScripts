/****************************************************
 * Wipes all events from a calendar between a given
 * start & end date
 * **************************************************/
function clearCalendar(){
  const CALENDAR_ID = '';
  let cal = CalendarApp.getCalendarById(CALENDAR_ID);
  let startDate = new Date('year','month','day');
  let endDate = new Date('year','month','day');
  let events = cal.getEvents(startDate,endDate);
  for (var i = 0;i<events.length;i++){
    var event = events[i];
    event.deleteEvent();
  }
}
