function onGmailMessageOpen(e) {
  try {
    if (!e || !e.gmail) {
      return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader().setTitle('メールを開いて実行してください'))
        .addSection(CardService.newCardSection().addWidget(
          CardService.newDecoratedText()
            .setTopLabel('ヒント')
            .setText('［デプロイ→テストとして導入→インストール］後、Gmailで実メールを開いて起動してください。')))
        .build();
    }

    GmailApp.setCurrentMessageAccessToken(e.gmail.accessToken);
    var msg = GmailApp.getMessageById(e.gmail.messageId);
    var subject = msg.getSubject();
    var threadLink = msg.getThread().getPermalink();

    // カレンダー候補（メイン/所有）
    var defCal = CalendarApp.getDefaultCalendar();
    var defId = defCal.getId();
    var all = CalendarApp.getAllCalendars();
    var cands = [];
    for (var i = 0; i < all.length; i++) {
      var c = all[i];
      if (c.isMyPrimaryCalendar() || c.isOwnedByMe()) cands.push(c);
    }
    if (cands.length === 0) cands = [defCal];

    var calSelect = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName('calId')
      .setTitle('作成先カレンダー');
    for (var j = 0; j < cands.length; j++) {
      var cal = cands[j];
      var label = cal.getName() + (cal.isMyPrimaryCalendar() ? '（メイン）' : '');
      calSelect.addItem(label, cal.getId(), cal.getId() === defId);
    }

    var title = CardService.newTextInput().setFieldName('title').setTitle('タイトル').setValue(subject);
    var start = CardService.newDateTimePicker()
      .setFieldName('start').setTitle('開始').setTimeZoneOffsetInMins(9 * 60)
      .setValueInMsSinceEpoch(new Date().getTime() + 15 * 60 * 1000);
    var end = CardService.newDateTimePicker()
      .setFieldName('end').setTitle('終了').setTimeZoneOffsetInMins(9 * 60)
      .setValueInMsSinceEpoch(new Date().getTime() + 45 * 60 * 1000);
    var desc = CardService.newTextInput()
      .setFieldName('desc').setTitle('説明').setMultiline(true)
      .setValue((msg.getPlainBody() || '').substring(0, 500) + '\n\n---\n' + threadLink);

    var action = CardService.newAction()
      .setFunctionName('createEvent') // 下記の"createEvent"関数を呼ぶ

    var btn = CardService.newTextButton()
      .setText('カレンダーに予定を作成')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(action);

    var section = CardService.newCardSection()
      .addWidget(calSelect).addWidget(title).addWidget(start).addWidget(end).addWidget(desc).addWidget(btn);

    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('このメールから予定を作成'))
      .addSection(section)
      .build();

  } catch (err) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('エラー'))
      .addSection(CardService.newCardSection().addWidget(
        CardService.newDecoratedText().setTopLabel('onGmailMessageOpen').setText((err && err.message) ? err.message : String(err))))
      .build();
  }
}

function createEvent(e) {
  try {
    /**
     * 予定作成処理
     */
    var inputs = (e && e.commonEventObject && e.commonEventObject.formInputs) || e.formInput || {};
    function _getStr(name) {
      var fi = inputs[name];
      if (fi && fi.stringInputs && fi.stringInputs.value && fi.stringInputs.value.length > 0) return fi.stringInputs.value[0];
      if (fi && fi.value && fi.value.length > 0) return fi.value[0];
      return '';
    }
    function _getMs(name, def) {
      var fi = inputs[name];
      if (fi && fi.dateTimeInput && typeof fi.dateTimeInput.msSinceEpoch === 'number') return fi.dateTimeInput.msSinceEpoch;
      return def;
    }

    var title = _getStr('title') || '（無題）';
    var desc = _getStr('desc') || '';
    var startMs = _getMs('start', new Date().getTime() + 15 * 60 * 1000);
    var endMs = _getMs('end', new Date().getTime() + 45 * 60 * 1000);
    var calId = _getStr('calId') || CalendarApp.getDefaultCalendar().getId();

    var cal = CalendarApp.getCalendarById(calId);
    cal.createEvent(title, new Date(startMs), new Date(endMs), { description: desc });


    /**
     * 予定作成後の表示
     */
    var done = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('予定作成完了'))
      .addSection(CardService.newCardSection().addWidget(
        CardService.newDecoratedText()
          .setText('予定を作成しました。Googleカレンダーに反映されているかご確認ください。')
          .setWrapText(true)))
      .build();

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(done))
      .build();

  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('エラー: ' + ((err && err.message) ? err.message : String(err))))
      .build();
  }
}
