
//Define calendar(s): addCalendar ("Unique Calendar Name", "Window title", "Form element's name", Form name")
addCalendar("StartDate", "::Calendar::", "startdatebox", "frm");
addCalendar("EndDate", "::Calendar::", "enddatebox", "frm");

// default settings for English
// Uncomment desired lines and modify its values
 setFont("arial",11);
 setWidth(90, 1, 20, 1);
 setColor("#f8f8f8", "#f8f8f8", "#f8f8f8", "#f8f8f8", "#cccccc", "#f8f8f8", "#cccccc");
 setFontColor("#ff0000","#ff0000", "#ff0000", "#ff0000", "#ff0000");
 setFormat("mm/dd/yyyy");
 setSize(200, 220, -200, 5);

// setWeekDay(0);
 setMonthNames("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
 setDayNames("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday");
 setLinkNames("[Close]", "[Clear]");
