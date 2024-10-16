(function () {
  "use strict";

  let messageBanner;

  // Office.js ライブラリを読み込みます。
  Office.onReady(function (reason) {
    $(() => {
      const element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();
      loadProps();
    });
  });

  // 添付ファイル名の一覧を作成します。
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      let returnString = "";
      
      for (let i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }

      return returnString;
    }

    return "None";
  }

  // 指定した名前、姓、およびメール アドレスで連絡先の詳細を書式設定します
    function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }

  // 書式設定された連絡先の詳細の一覧を構築します。$loc_script_mail_commands_read_js_comment3$)$$  function buildEmailAddressesString(addresses) {    if (addresses && addresses.length > 0) {      let returnString = "";       for (let i = 0; i < addresses.length; i++) {        if (i > 0) {          returnString = returnString + "<br/>";        }        returnString = returnString + buildEmailAddressString(addresses[i]);      }       return returnString;    }     return "None";  }   // $$LOC(アイテムのベース オブジェクトからプロパティを読み込んだ後、
  // メッセージ固有のプロパティを読み込みます。
  function loadProps() {
    const item = Office.context.mailbox.item;

    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    $('#message-props').show();

    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.from));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
    $('#to').html(buildEmailAddressesString(item.to));
  }

  // 通知を表示するヘルパー関数
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();