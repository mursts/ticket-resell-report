import GmailLabel = GoogleAppsScript.Gmail.GmailLabel;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
declare let global: any;

const searchQuery = '【Ｊリーグチケット】リセールチケット売買成立のお知らせ has:nouserlabels';
const regexSearch =
  '公演名：(.*)[\\s\\S]*公演日：([0-9]{4}/[0-9]{2}/[0-9]{2})[\\s\\S]*送金予定金額：(.*)円';
const gmailLabel = 'resell';

const getLabel = (): GmailLabel => {
  return GmailApp.getUserLabelByName(gmailLabel);
};

global.run = (): void => {
  const sheet = SpreadsheetApp.getActive();
  GmailApp.search(searchQuery).forEach(thread => {
    thread.getMessages().forEach(message => {
      const body = message.getPlainBody();
      const m = body.match(regexSearch);
      if (!m) {
        return;
      }
      const title = m[1];
      const date = m[2];
      const amount = m[3].replace(',', '');
      const issueDate = Utilities.formatDate(message.getDate(), 'JST', 'yyyy/MM/dd');
      sheet.appendRow([issueDate, date, title, amount]);

      // add label
      let label = getLabel();
      if (label === null) {
        label = GmailApp.createLabel(gmailLabel);
      }
      label.addToThread(thread);
    });
  });
};
