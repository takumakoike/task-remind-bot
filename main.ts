declare const ChatWorkClient: {
    factory(options: { token: string }): any;
};

interface ChatworkTask {
    room: { name: string; room_id: string };
    assigned_by_account: { name: string };
    body: string;
    message_id: string;
    limit_time: number;
}

function alert_my_task(): void {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('my_tasks');
    if (!sheet) return;

    const props    = PropertiesService.getScriptProperties();
    const token    = props.getProperty("CHATWORK_TOKEN") ?? Config.CHATWORK_TOKEN;
    const roomId   = props.getProperty("ROOM_ID")        ?? Config.ROOM_ID;
    const botToken = props.getProperty("BOT_TOKEN")      ?? Config.BOT_TOKEN;

    const fetchParam: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        headers: { "X-ChatWorkToken": token },
        method: "get",
        muteHttpExceptions: true
    };

    const res = UrlFetchApp.fetch("https://api.chatwork.com/v2/my/tasks", fetchParam).getContentText();
    const cw  = ChatWorkClient.factory({ token: botToken });

    if (!res || res.length < 1) {
        Browser.msgBox("未完了タスクはございませんでした。");
        cw.sendMessage({ room_id: roomId, body: "未完了タスクはございませんでした。" });
        return;
    }

    const tasks: ChatworkTask[] = JSON.parse(res);
    const writeData: string[][] = [];
    const openTaskMessages: string[] = [];

    tasks.forEach((task, i) => {
        const limit    = Utilities.formatDate(new Date(task.limit_time * 1000), 'Asia/Tokyo', 'yyyy年MM月dd日');
        const taskLink = `https://www.chatwork.com/#!rid${task.room.room_id}-${task.message_id}`;

        writeData.push([limit, task.room.name, task.assigned_by_account.name, task.body]);
        openTaskMessages.push(
            `＝＝＝${i + 1}件目＝＝＝\n【期日】：${limit}\n【グループチャット名】：${task.room.name}\n【タスク詳細】\n[info]${task.body}[/info]\n【メッセージリンク】\n${taskLink}\n`
        );
    });

    sheet.getRange(6, 1, tasks.length, 4).clearContent().setValues(writeData);

    const date    = new Date();
    const memId   = sheet.getRange(1, 2).getValue() as string;
    const name    = sheet.getRange(1, 3).getValue() as string;
    const msgBody = `[To:${memId}]${name}さん\n${date.getMonth() + 1}月${date.getDate()}日現在、未完了タスクは${tasks.length}件ございます。\n\n${openTaskMessages.join("\n[hr]\n\n")}\n以上です。`;

    cw.sendMessage({ room_id: roomId, body: msgBody });

    const triggerTargetId = sheet.getRange(4, 4).getValue() as string;
    ScriptApp.getProjectTriggers().forEach(trigger => {
        if (trigger.getUniqueId() === triggerTargetId) {
            ScriptApp.deleteTrigger(trigger);
        }
    });
}

function setTrigger(timeCol: number, idCol: number): void {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('my_tasks');
    if (!sheet) return;

    const remHour = sheet.getRange(4, timeCol).getValue();
    if (remHour === "") {
        Browser.msgBox(`${timeCol === 2 ? "B" : "C"}4セルに時刻を設定してください`);
        return;
    }
    const trigger = ScriptApp.newTrigger('alert_my_task')
        .timeBased()
        .atHour(remHour)
        .everyDays(1)
        .inTimezone('Asia/Tokyo')
        .create();
    sheet.getRange(4, idCol).setValue(trigger.getUniqueId());
}

function set_trigger()        { setTrigger(2, 4); }
function set_trigger_second() { setTrigger(3, 5); }