function collectCalendarEvents() {
  const ui = SpreadsheetApp.getUi();
  ui.alert("✅ 集計を開始します。完了後、別シートに出力されます。");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("設定");
  const configData = configSheet.getDataRange().getValues();

  const today = new Date();
  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");

  const sheetNames = ss.getSheets().map(sheet => sheet.getName());
  const regex = new RegExp(`^${dateStr}_(\\d+)$`);
  let maxIndex = 0;
  sheetNames.forEach(name => {
    const match = name.match(regex);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxIndex) maxIndex = num;
    }
  });

  const newSheetName = `${dateStr}_${maxIndex + 1}`;
  const outputSheet = ss.insertSheet(newSheetName);
  outputSheet.appendRow(["名前", "日付", "時間帯", "タイトル", "工数"]);

  for (let i = 1; i < configData.length; i++) {
    const [name, email, fromDateStr, toDateStr, target] = configData[i];

    if (target !== 0) continue;

    const fromDate = new Date(fromDateStr);
    const toDate = new Date(toDateStr);
    toDate.setDate(toDate.getDate() + 1);

    const calendar = CalendarApp.getCalendarById(email);
    if (!calendar) {
      Logger.log(`⚠ カレンダーが見つかりません: ${email}`);
      continue;
    }

    const events = calendar.getEvents(fromDate, toDate);

    events.forEach(event => {
      const start = event.getStartTime();
      const end = event.getEndTime();
      const duration = (end - start) / (1000 * 60 * 60);
      const durationRounded = Math.floor(duration * 4) / 4;

      outputSheet.appendRow([
        name,
        Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy/MM/dd"),
        Utilities.formatDate(start, Session.getScriptTimeZone(), "HH:mm") + "〜" +
        Utilities.formatDate(end, Session.getScriptTimeZone(), "HH:mm"),
        event.getTitle(),
        durationRounded + "h"
      ]);
    });
  }

  SpreadsheetApp.flush();
  ui.alert("✅ 集計が完了しました！");
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ カレンダーの集計が完了しました！");
}



