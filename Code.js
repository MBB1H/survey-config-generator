/**
==========================================
Changelog
==========================================

v0.1.0 — 2026-07-18
- Создан каркас генератора.
- Реализовано чтение Excel.
- Добавлены модели Survey и Question.
- Реализовано определение вопросов.
- Добавлены парсеры radio, checkbox,
  numberInput и textArea.
- Реализован TransitionParser.

v0.2.0 — 2026-07-19
- Добавлена генерация итогового JSON.
- Реализованы обязательные и необязательные вопросы.
- Добавлены collectedKey и isCollected.
- Добавлены defaultNextId и описание вопросов.

v0.3.0 — 2026-07-20
- Добавлена поддержка numberInput.
- Реализовано определение возраста.
- Добавлены валюты и safeMaxThreshold.
- Добавлено определение диапазонов из заметок.

v0.4.0 — 2026-07-22
- Реализован SurveyFlowCalculator.
- Добавлен расчет blockWeight.
- Добавлен расчет progressWeight.
- Добавлена проверка циклических переходов.

v0.5.0 — 2026-07-24
- Добавлены условные переходы checkbox.
- Реализованы nextIdPrecalculate.
- Реализованы progressWeightPrecalculate.
- Добавлена поддержка составных условий.

v0.6.0 — 2026-07-25
- Добавлена поддержка otherOption.
- Добавлена генерация карточки результата.
- Добавлена автоматическая сборка списка статей.

v0.7.0 — 2026-07-26
- Добавлена поддержка select.
- Добавлена поддержка otherOption для select.
- Обновлен QuestionValidator.

v0.8.0 — 2026-07-27
- Добавлена типографика.
- Добаваны неразрывные пробелы.
- Добавлено форматирование диапазонов.
- Добавлена обработка частиц и постфиксов.

v0.9.0 — 2026-07-29
- Добавлена поддержка скрытых листов Excel.
- Реализован автоматический выбор рабочего листа.
- Добавлена поддержка листов «Вопросы2», «Вопросы3» и аналогичных.

v1.0.0 — 2026-08-04
- Проведено тестирование на 20+ реальных продовых опросах.
- Проверены все поддерживаемые типы вопросов.
- Проверены сложные маршруты и расчеты весов.
- Добавлены диагностические сообщения для неподдерживаемых конструкций.
- Генератор признан готовым к использованию редакцией.
 */

function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('Генераторы')
    .addItem(
      '📊 Генератор опросов',
      'openSurveySidebar'
    )
    .addToUi();
}


function openSurveySidebar() {
  const html =
    HtmlService
      .createTemplateFromFile(
        'Sidebar'
      )
      .evaluate()
      .setTitle(
        'Генератор опросов'
      );

  SpreadsheetApp
    .getUi()
    .showSidebar(
      html
    );
}


function processWorkbook(
  workbookData,
  surveySlug,
  resultHeading,
  resultDescription,
  resultArticles
) {
  const model =
    SurveyBuilder.build(
      workbookData
    );

  model.metadata =
    model.metadata &&
    typeof model.metadata === 'object'
      ? model.metadata
      : {};

  Object.assign(
    model.metadata,
    {
      surveySlug:
        String(
          surveySlug || ''
        ).trim(),

      resultHeading:
        String(
          resultHeading || ''
        ).trim(),

      resultDescription:
        String(
          resultDescription || ''
        ).trim(),

      resultArticles:
        Array.isArray(
          resultArticles
        )
          ? resultArticles
          : []
    }
  );

  model.json =
    JsonBuilder.build(
      model
    );

  model.jsonText =
    JSON.stringify(
      model.json,
      null,
      2
    );

  return model;
}