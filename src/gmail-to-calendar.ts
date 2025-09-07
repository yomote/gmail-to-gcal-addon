/**
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * Gmail拡張機能のメッセージ開封時のイベントハンドラー
 * メールの内容からGoogleカレンダーの予定作成フォームを表示する
 */
export async function onGmailMessageOpen(
  e: GoogleAppsScript.Addons.EventObject
): Promise<GoogleAppsScript.Card_Service.Card> {
  try {
    if (!e || !e.gmail) {
      return CardService.newCardBuilder()
        .setHeader(
          CardService.newCardHeader().setTitle('メールを開いて実行してください')
        )
        .addSection(
          CardService.newCardSection().addWidget(
            CardService.newDecoratedText()
              .setTopLabel('ヒント')
              .setText(
                '［デプロイ→テストとして導入→インストール］後、Gmailで実メールを開いて起動してください。'
              )
          )
        )
        .build();
    }

    GmailApp.setCurrentMessageAccessToken(e.gmail.accessToken);
    const msg = GmailApp.getMessageById(e.gmail.messageId);
    const subject = msg.getSubject();
    const threadLink = msg.getThread().getPermalink();

    // カレンダー候補（メイン/所有）
    const defCal = CalendarApp.getDefaultCalendar();
    const defId = defCal.getId();
    const all = CalendarApp.getAllCalendars();
    let cands: GoogleAppsScript.Calendar.Calendar[] = [];

    for (let i = 0; i < all.length; i++) {
      const c = all[i];
      if (c.isMyPrimaryCalendar() || c.isOwnedByMe()) {
        cands.push(c);
      }
    }
    if (cands.length === 0) {
      cands = [defCal];
    }

    const calSelect = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName('calId')
      .setTitle('作成先カレンダー');

    for (let j = 0; j < cands.length; j++) {
      const cal = cands[j];
      const label =
        cal.getName() + (cal.isMyPrimaryCalendar() ? '（メイン）' : '');
      calSelect.addItem(label, cal.getId(), cal.getId() === defId);
    }

    const title = CardService.newTextInput()
      .setFieldName('title')
      .setTitle('タイトル')
      .setValue(subject);

    // メール本文から日時を抽出（OpenAI APIを利用）
    const extracted = await extractDateTime(msg.getPlainBody());
    const startDate =
      extracted?.startMs ?? new Date().getTime() + 15 * 60 * 1000;
    const endDate = extracted?.endMs ?? new Date().getTime() + 45 * 60 * 1000;

    const start = CardService.newDateTimePicker()
      .setFieldName('start')
      .setTitle('開始')
      .setTimeZoneOffsetInMins(9 * 60) // 日本時間
      .setValueInMsSinceEpoch(startDate);

    const end = CardService.newDateTimePicker()
      .setFieldName('end')
      .setTitle('終了')
      .setTimeZoneOffsetInMins(9 * 60) // 日本時間
      .setValueInMsSinceEpoch(endDate);

    const desc = CardService.newTextInput()
      .setFieldName('desc')
      .setTitle('説明')
      .setMultiline(true)
      .setValue(
        (msg.getPlainBody() || '').substring(0, 500) + '\n\n---\n' + threadLink
      );

    const action = CardService.newAction().setFunctionName('createEvent'); // 下記の"createEvent"関数を呼ぶ

    const btn = CardService.newTextButton()
      .setText('カレンダーに予定を作成')
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(action);

    const section = CardService.newCardSection()
      .addWidget(calSelect)
      .addWidget(title)
      .addWidget(start)
      .addWidget(end)
      .addWidget(desc)
      .addWidget(btn);

    return CardService.newCardBuilder()
      .setHeader(
        CardService.newCardHeader().setTitle('このメールから予定を作成')
      )
      .addSection(section)
      .build();
  } catch (err) {
    return CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('エラー'))
      .addSection(
        CardService.newCardSection().addWidget(
          CardService.newDecoratedText()
            .setTopLabel('onGmailMessageOpen')
            .setText(
              err && (err as Error).message
                ? (err as Error).message
                : String(err)
            )
        )
      )
      .build();
  }
}

/**
 * カレンダー予定作成処理
 * フォームから送信された情報を使ってGoogleカレンダーに予定を作成する
 */
export function createEvent(
  e: GoogleAppsScript.Addons.EventObject
): GoogleAppsScript.Card_Service.ActionResponse {
  try {
    /**
     * 予定作成処理
     */
    const inputs =
      (e && e.commonEventObject && e.commonEventObject.formInputs) || {};

    const title = getStringFromFormInput(inputs, 'title') || '（無題）';
    const desc = getStringFromFormInput(inputs, 'desc') || '';
    const startMs = getMillisecondsFromFormInput(
      inputs,
      'start',
      new Date().getTime() + 15 * 60 * 1000
    );
    const endMs = getMillisecondsFromFormInput(
      inputs,
      'end',
      new Date().getTime() + 45 * 60 * 1000
    );

    const calId =
      getStringFromFormInput(inputs, 'calId') ||
      CalendarApp.getDefaultCalendar().getId();
    const cal = CalendarApp.getCalendarById(calId);
    cal.createEvent(title, new Date(startMs), new Date(endMs), {
      description: desc,
    });

    /**
     * 予定作成後の表示
     */
    const done = CardService.newCardBuilder()
      .setHeader(CardService.newCardHeader().setTitle('予定作成完了'))
      .addSection(
        CardService.newCardSection().addWidget(
          CardService.newDecoratedText()
            .setText(
              '予定を作成しました。Googleカレンダーに反映されているかご確認ください。'
            )
            .setWrapText(true)
        )
      )
      .build();

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(done))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(
        CardService.newNotification().setText(
          'エラー: ' +
            (err && (err as Error).message
              ? (err as Error).message
              : String(err))
        )
      )
      .build();
  }
}

/**
 * メール文面から日時を抽出する
 *
 * OpenAI APIを利用して、メール本文から予定の開始・終了日時を解析する
 */
async function extractDateTime(
  emailBody: string
): Promise<{ startMs: number; endMs: number } | null> {
  try {
    // OpenAI API keyを取得（PropertiesServiceから取得するか、直接設定）
    const apiKey =
      PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
      console.log('OpenAI API key not found. Skipping AI extraction.');
      return null;
    }

    const prompt = `
    以下のメール本文から予定の開始日時と終了日時を抽出してください。

    メール本文:
    ${emailBody}
    
    出力形式（JSONのみ）:
    { "startDate": YYYY-MM-DDTHH:mm:ssZ, "endDate": YYYY-MM-DDTHH:mm:ssZ }

    日時が見つからない場合は各値に"Not Found"を設定してください。
    `;

    const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify({
        model: 'gpt-4.1',
        input: prompt,
        text: {
          format: {
            type: 'json_schema',
            name: 'extracted_dates',
            schema: {
              type: 'object',
              properties: {
                startDate: { type: 'string' },
                endDate: { type: 'string' },
              },
              required: ['startDate', 'endDate'],
              additionalProperties: false,
            },
            strict: true,
          },
        },
      }),
    });

    const responseData = JSON.parse(response.getContentText());
    console.log('OpenAI full response:', responseData);

    const content = responseData.output[0].content[0].text;

    if (content) {
      console.log('OpenAI response:', content);
      const parsedData = JSON.parse(content);
      if (
        parsedData &&
        typeof parsedData.startDate === 'string' &&
        typeof parsedData.endDate === 'string'
      ) {
        if (
          parsedData.startDate === 'Not Found' ||
          parsedData.endDate === 'Not Found'
        ) {
          return null;
        }

        return {
          startMs: Date.parse(parsedData.startDate),
          endMs: Date.parse(parsedData.endDate),
        };
      }
    }
  } catch (err) {
    console.error('Error extracting datetime with OpenAI:', err);
  }

  return null;
}

/**
 * フォーム入力から文字列値を取得するヘルパー関数
 */
function getStringFromFormInput(
  inputs: GoogleAppsScript.Addons.CommonEventObject['formInputs'],
  name: string
): string {
  const fi = inputs?.[name];
  if (
    fi &&
    fi.stringInputs &&
    fi.stringInputs.value &&
    fi.stringInputs.value.length > 0
  ) {
    return fi.stringInputs.value[0];
  }
  return '';
}

/**
 * フォーム入力から日時のミリ秒値を取得するヘルパー関数
 */
function getMillisecondsFromFormInput(
  inputs: GoogleAppsScript.Addons.CommonEventObject['formInputs'],
  name: string,
  def: number
): number {
  const fi = inputs?.[name];
  if (
    fi &&
    fi.dateTimeInput &&
    typeof fi.dateTimeInput.msSinceEpoch === 'number'
  ) {
    return fi.dateTimeInput.msSinceEpoch;
  }
  return def;
}
