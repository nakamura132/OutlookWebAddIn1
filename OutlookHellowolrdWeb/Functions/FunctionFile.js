// Office.js ライブラリを読み込みます。
Office.onReady();

// 通知バーにステータス メッセージを追加するヘルパー関数。
function statusUpdate(icon, text, event) {
  const details = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: icon,
    message: text,
    persistent: false
  };
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", details, { asyncContext: event }, asyncResult => {
    const event = asyncResult.asyncContext;
    event.completed();
  });
}
// 通知バーを表示します。
function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!", event);
}

// マニフェストで指定された関数名を、対応する JavaScript にマッピングします。
Office.actions.associate("defaultStatus", defaultStatus);