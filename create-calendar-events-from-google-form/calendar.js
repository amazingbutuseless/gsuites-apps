/**
 * ACTION REQIRED
 * @brief Google Spredsheets에 사용자 메뉴명으로 표시될 이름
 * */ 
const MENU_NAME = 'iPF 근태 캘린더';

/**
 * ACTION REQUIRED
 * @brief 데이터를 가져올 시트명. 설문지의 답변이 기록되는 시트명을 기록
 */
const SHEET_NAME = 'Form Responses 1';

/**
 * ACTION REQUIRED
 * @brief 생성할 캘린더 이름
 */
const CALENDAR_NAME = '근태';

const SATURDAY = 6;
const SUNDAY = 0;
const timezoneOffset = 540 * 60 * 1000; // GMT+9:00
const oneDay = 60 * 60 * 24 * 1000;


/**
 * @brief 답변 내용 중 사유와 기간을 이용하여 캘린더에 표시될 제목 중 휴가 종류를 결정, 반환
 * @param (enum) reason 
 * @param (string) duration 
 */
function getLeaveType(reason, duration) {
  let type = '';

  switch (reason) {
    case '재택근무':
    case '외부 교육':
      type = reason;
      break;

    case '출장(1일 이상)':
      type = '출장';
      break;

    default:
      if (duration.startsWith('0.5')) {
        const details = /0\.5일\((오전|오후)\)/.exec(duration);
        type = `${details[1]}반차`;

      } else {
        type = '휴가';
      }
  }

  return type;
}


/**
 * @brief 설문지 작성자, 사유/기간을 바탕으로 결정된 캘린더 이벤트 제목을 반환
 * @param (string) name 
 * @param (string) reason 
 * @param (string) duration 
 */
function generateEventTitle(name, reason, duration) {
  const type = getLeaveType(reason, duration + '');
  return `${type}(${name})`;
}


/**
 * @brief 사용자가 질문지를 submit 했을 때 발생하는 이벤트를 처리하는 핸들러, 사용자 답변을 바탕으로 캘린더에 이벤트를 등록함
 * @param (FormEvent) e 
 */
function onFormSubmit(e) {
  const name = e.namedValues['본인 이름을 적어주세요.'][0];
  const reason = e.namedValues['부재 사유는 무엇입니까?'][0];
  const duration = e.namedValues['부재 일수를 적어주세요.(업무일 기준)'][0];

  const start = e.namedValues['부재 시작일은 언제입니까?'][0];
  const end = e.namedValues['부재 종료일은 언제입니까? (2일 이상의 부재인 경우만 기재)'][0];

  const calId = ScriptProperties.getProperty('calId');
  const cal = CalendarApp.getCalendarById(calId);

  addEventToCalendar(
    cal,
    generateEventTitle(name, reason, duration),
    start,
    end
  );
}


/**
 * @brief 신청 기간 내 주말이 포함되어 있을 경우, 주말을 제거된 복수개의 이벤트 등록 할 수 있도록 기간을 분리하여 반환
 * @param (Date) startDate 
 * @param (Date) endDate 
 */
function getEventsWithNoWeekend(startDate, endDate) {
  let events = [];

  let eventIdx = 0;

  let startTimestamp = startDate.getTime();

  while (startTimestamp < endDate.getTime()) {
    const day = new Date(startTimestamp).getDay();

    if (day === SATURDAY) {
      events[eventIdx]['end'] = new Date(startTimestamp - 1000);
      eventIdx++;

    } else if (day !== SUNDAY && (events.length === 0 || events[events.length - 1].hasOwnProperty('end'))) {
      events.push({
        start: new Date(startTimestamp),
      });

    }

    startTimestamp += oneDay;
  }

  events[eventIdx]['end'] = endDate;

  return events;
}


/**
 * @brief 이벤트를 Calendar에 등록
 * @param (CalendarApp) cal 
 * @param (string) title 
 * @param (string) start 
 * @param (string) end 
 */
function addEventToCalendar(cal, title, start, end) {
  let events = [];

  const startDate = new Date(new Date(start).getTime() + timezoneOffset);
  let endDate = end.length < 1 ? undefined : new Date(end);

  if (endDate) {
    endDate = new Date(endDate.getTime() + (oneDay - 1000) + timezoneOffset);

    events = getEventsWithNoWeekend(startDate, endDate);
  } else {
    events.push({
      start: startDate,
      end: endDate,
    });
  }

  events.forEach((event) => {
    cal.createAllDayEvent(title, event.start, event.end);
  });
}


/**
 * @brief 캘린더 초기 세팅. 캘린더를 생성하고, 시트에 작성되어 있던 내용을 일괄로 캘린더에 등록
 * @param (*) values 
 */
function setUpCalendar(values) {
  const cal = CalendarApp.createCalendar(CALENDAR_NAME);
  cal.setTimeZone('Asia/Seoul');

  values.shift();

  values.forEach((value) => {
    const name = value[1];
    const reason = value[2];
    const duration = value[3];

    const start = value[4];
    const end = value[5];

    addEventToCalendar(
      cal,
      generateEventTitle(name, reason, duration),
      start,
      end
    );
  });

  ScriptProperties.setProperty('calId', cal.getId());
}


/**
 * @brief AppScript 초기 설정
 */
function init() {
  if (ScriptProperties.getProperty('calId')) {
    Browser.msgBox(`${ MENU_NAME } has already set up.`);
  }

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  var range = sheet.getDataRange();
  var values = range.getValues();
  setUpCalendar(values);

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  ss.removeMenu(MENU_NAME);
}


function onOpen() {
  const menu = [
    {
      name: 'Set Up',
      functionName: 'init',
    },
  ];

  SpreadsheetApp.getActive().addMenu(MENU_NAME, menu);
}
