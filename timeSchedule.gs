function createTimeSchedule() {　　　　
	Browser.msgBox("スケジュールの作成を開始します。");

	var sheet = SpreadsheetApp.getActive().getSheetByName("スケジュール");
	var range = sheet.getRange("B3:O50");

	initSheet_(range);

	var cells;

	cells = sheet.getRange(1, 2, 1, 16);
	var dateHeaders = cells.getValues();

	cells = sheet.getRange(3, 1, 54, 1);
	var timeHeaders = cells.getValues();

	var headers = [];
	headers = createHeaders_(headers, dateHeaders, timeHeaders);

	var schedules = [];
	var events = getCalendar_();
	schedules = createSchedules_(schedules, events, headers);

	insertSchedules_(sheet, schedules);

	Browser.msgBox("スケジュールの作成が完了しました。");
}

function initSheet_(range) {
	range.clear();
	range.setBackground("#FFFFFF");

	range.setHorizontalAlignment("center");
	range.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.DASHED);
	range.setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function getCalendar_() {
	var date = new Date();
	var startDate　 = 　new Date(date);
	var endDate　 = 　new Date(date);
	endDate.setMonth(endDate.getMonth()　 + 　1);

	var user = Session.getActiveUser();
	var calender　 = 　CalendarApp.getCalendarById(user.getEmail());

	return calender.getEvents(startDate, endDate);
}

function createHeaders_(headers, dateHeaders, timeHeaders) {
	var date;
	var time;

	for (var i = 0; i < dateHeaders[0].length; i++) {
		if (dateHeaders[0][i] != "") {
			for (var j = 0; j < timeHeaders.length; j++) {
				if (timeHeaders[j] != "") {
					date = formatDateTime_(new Date(dateHeaders[0][i]), "yyyy/MM/dd");
					time = formatDateTime_(new Date(timeHeaders[j]), "HH:mm:ss");
					headers.push(date + " " + time + "-" + String(i + 2) + ":" + String(j + 3));
				}
			}
		}
	}

	return headers;
}

function formatDateTime_(datetime, format) {
	format = format.replace(/yyyy/g, datetime.getFullYear());
	format = format.replace(/MM/g, ('0' + (datetime.getMonth() + 1)).slice(-2));
	format = format.replace(/dd/g, ('0' + datetime.getDate()).slice(-2));
	format = format.replace(/HH/g, ('0' + datetime.getHours()).slice(-2));
	format = format.replace(/mm/g, ('0' + datetime.getMinutes()).slice(-2));
	format = format.replace(/ss/g, ('0' + datetime.getSeconds()).slice(-2));
	format = format.replace(/SSS/g, ('00' + datetime.getMilliseconds()).slice(-3));

	return format;
}

function createSchedules_(schedules, events, headers) {
	var startDateTime;
	var startYear;
	var startMonth;
	var startDay;
	var startHour;
	var startMinute;
	var startTime;

	var endDateTime;
	var endYear;
	var endMonth;
	var endDay;
	var endHour;
	var endMinute;
	var endTime;

	var block;
	var arrayHeader;

	for each(var event in events) {
		startDateTime　 = event.getStartTime();

		startYear = formatDateTime_(startDateTime, "yyyy");
		startMonth = formatDateTime_(startDateTime, "MM");
		startDay = formatDateTime_(startDateTime, "dd");
		startHour = formatDateTime_(startDateTime, "HH");
		startMinute = formatDateTime_(startDateTime, "mm");

		startTime = parseDateTime_(startHour, startMinute, "00", "30");
		startDateTime = concatDatetime_(startYear, startMonth, startDay, startTime);

		endDateTime　 = event.getEndTime();

		endYear = formatDateTime_(endDateTime, "yyyy");
		endMonth = formatDateTime_(endDateTime, "MM");
		endDay = formatDateTime_(endDateTime, "dd");
		endHour = formatDateTime_(endDateTime, "HH");
		endMinute = formatDateTime_(endDateTime, "mm");

		endTime = parseDateTime_(endDateTime.getHours(), endMinute, "30", "00");
		endDateTime = concatDatetime_(endYear, endMonth, endDay, endTime);

		block = ((((endDateTime - startDateTime) / 1000) / 3600) / 0.5);

		for each(var header in headers) {
			arrayHeader = header.split("-");

			if (arrayHeader[0] == formatDateTime_(startDateTime, "yyyy/MM/dd HH:mm:ss")) {
				schedules.push(arrayHeader[1] + "-" + block + "-" + event.getTitle());
			}
		}
	}

	return schedules;
}

function parseDateTime_(hour, minute, minMinute, maxMinute) {
	if (parseFloat(minute) != "00" && parseFloat(minute) != "30") {
		if (parseFloat(minute) < 30) {
			minute = minMinute;
		} else if (parseFloat(minute) < 59) {
			if (maxMinute == "00") {
				hour = hour + 1;
			}

			minute = maxMinute;
		}
	}

	return hour + ":" + minute + ":00";
}

function concatDatetime_(year, month, date, time) {
	return new Date(year + "/" + month + "/" + date + " " + time);
}

function insertSchedules_(sheet, schedules) {
	var arraySchedule;

	var cell;
	var arrayCell;
	var column;
	var row;

	var block;
	var title;

	var lastRow;
	var lastRange;
	var lastCell;

	for each(var schedule in schedules) {
		arraySchedule = schedule.split("-");

		cell = arraySchedule[0];
		arrayCell = cell.split(":");
		column = arrayCell[0];
		row = arrayCell[1];

		block = arraySchedule[1];
		title = arraySchedule[2];

		for (var i = 0; i < block; i++) {
			lastRow = parseInt(row) + parseInt(i);
			lastRange = sheet.getRange(lastRow, column);
			lastCell = lastRange.getValue();

			if (lastCell != "") {
				lastRange.setValue(lastCell + "\r\n" + title);
				lastRange.setBackground("#FF0000");
			} else {
				lastRange.setValue(title);
			}
		}
	}
