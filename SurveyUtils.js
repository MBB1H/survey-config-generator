/**
 * ==========================================
 * 1. МОДЕЛИ ДАННЫХ
 * ==========================================
 *
 * Внутренние структуры, в которых хранятся
 * данные опроса, вопросов и вариантов ответа.
 */


/**
 * Внутренняя модель опроса.
 */
class SurveyModel {

  constructor() {
    this.metadata = {};
    this.sheets = [];
    this.questions = [];
    this.results = [];
    this.errors = [];
    this.questionBlocks = [];
    this.warnings = [];
  }

}


/**
 * Внутренняя модель вопроса.
 */
class Question {

  constructor() {
    this.errors = [];
    this.id = '';
    this.number = 0;
    this.title = '';
    this.description = '';
    this.type = '';
    this.options = [];

    this.transition = null;
    this.note = '';

    this.defaultNextId = null;
    this.progressWeight = 1;
    this.blockWeight = null;
  }

}


/**
 * Внутренняя модель варианта ответа.
 */
class Option {

  constructor() {
    this.text = '';
    this.transition = null;
    this.note = '';
    this.progressWeight = null;
  }

}


/**
 * ==========================================
 * КОНЕЦ РАЗДЕЛА: МОДЕЛИ ДАННЫХ
 * ==========================================
 */


/**
 * ==========================================
 * 2. ЧТЕНИЕ И НОРМАЛИЗАЦИЯ EXCEL
 * ==========================================
 *
 * Проверяет входные данные Excel,
 * нормализует листы, строки, значения ячеек
 * и информацию об объединенных диапазонах.
 */


class WorkbookReader {

  static read(workbookData) {
    this.validateWorkbook(workbookData);

    const model = new SurveyModel();

    model.metadata = {
      fileName:
        String(
          workbookData.fileName || ''
        ),

      sheetCount:
        workbookData.sheets.length
    };

    model.sheets =
      workbookData.sheets.map(
        (sheet, index) =>
          this.normalizeSheet(
            sheet,
            index
          )
      );

    return model;
  }

  static validateWorkbook(workbookData) {
    if (
      !workbookData ||
      typeof workbookData !== 'object'
    ) {
      throw new Error(
        'Получены некорректные данные Excel.'
      );
    }

    if (
      !Array.isArray(
        workbookData.sheets
      )
    ) {
      throw new Error(
        'В Excel отсутствует список листов.'
      );
    }

    if (
      workbookData.sheets.length === 0
    ) {
      throw new Error(
        'В Excel не найдено ни одного листа.'
      );
    }
  }

  static normalizeSheet(
    sheet,
    index
  ) {
    const name =
      String(
        sheet.name || ''
      ).trim();

    const rows =
      this.normalizeRows(
        sheet.rows
      );

    const columnCount =
      this.getColumnCount(
        rows
      );

    return {
      index,
      name,

      normalizedName:
        this.normalizeName(
          name
        ),

      type:
        this.detectSheetType(
          name
        ),

      isHidden:
        Boolean(
          sheet.isHidden
        ),

      rowCount:
        rows.length,

      columnCount,

      rows,

      merges:
        this.normalizeMerges(
          sheet.merges
        )
    };
  }

  static normalizeRows(
    sourceRows
  ) {
    if (
      !Array.isArray(
        sourceRows
      )
    ) {
      return [];
    }

    const rows =
      sourceRows.map(
        row => {
          if (
            !Array.isArray(
              row
            )
          ) {
            return [];
          }

          return row.map(
            cell =>
              this.normalizeCellValue(
                cell
              )
          );
        }
      );

    return this.removeTrailingEmptyRows(
      rows
    );
  }

  static normalizeCellValue(
    value
  ) {
    if (
      value === null ||
      value === undefined
    ) {
      return '';
    }

    if (
      typeof value === 'string'
    ) {
      return value
        .replace(
          /\r\n?/g,
          '\n'
        )
        .replace(
          /\u00A0/g,
          ' '
        )
        .trim();
    }

    return String(
      value
    ).trim();
  }

  static removeTrailingEmptyRows(
    rows
  ) {
    let lastFilledRowIndex =
      rows.length - 1;

    while (
      lastFilledRowIndex >= 0 &&
      this.isEmptyRow(
        rows[
          lastFilledRowIndex
        ]
      )
    ) {
      lastFilledRowIndex--;
    }

    return rows.slice(
      0,
      lastFilledRowIndex + 1
    );
  }

  static isEmptyRow(
    row
  ) {
    return row.every(
      cell =>
        cell === ''
    );
  }

  static getColumnCount(
    rows
  ) {
    return rows.reduce(
      (
        maximum,
        row
      ) =>
        Math.max(
          maximum,
          row.length
        ),
      0
    );
  }

  static normalizeMerges(
    sourceMerges
  ) {
    if (
      !Array.isArray(
        sourceMerges
      )
    ) {
      return [];
    }

    return sourceMerges
      .filter(
        merge =>
          merge &&
          Number.isInteger(
            merge.startRow
          ) &&
          Number.isInteger(
            merge.startColumn
          ) &&
          Number.isInteger(
            merge.endRow
          ) &&
          Number.isInteger(
            merge.endColumn
          )
      )
      .map(
        merge => ({
          startRow:
            merge.startRow,

          startColumn:
            merge.startColumn,

          endRow:
            merge.endRow,

          endColumn:
            merge.endColumn
        })
      );
  }

  static detectSheetType(
    name
  ) {
    const normalizedName =
      this.normalizeName(
        name
      );

    if (
      /^_?вопросы\d*$/.test(
        normalizedName
      )
    ) {
      return 'questions';
    }

    if (
      normalizedName ===
        '_типы вопросов' ||
      normalizedName ===
        'типы вопросов'
    ) {
      return 'questionTypes';
    }

    return 'unknown';
  }

  static normalizeName(
    value
  ) {
    return String(
      value || ''
    )
      .trim()
      .toLowerCase()
      .replace(
        /ё/g,
        'е'
      )
      .replace(
        /\s+/g,
        ' '
      );
  }

}


/**
 * ==========================================
 * КОНЕЦ РАЗДЕЛА:
 * ЧТЕНИЕ И НОРМАЛИЗАЦИЯ EXCEL
 * ==========================================
 */


/**
 * ==========================================
 * 3. ОПРЕДЕЛЕНИЕ БЛОКОВ ВОПРОСОВ
 * ==========================================
 *
 * Находит начало каждого вопроса
 * и собирает относящиеся к нему строки.
 */


class QuestionDetector {

  static detect(model) {
    const questionsSheetIndex =
      model.metadata
        .questionsSheetIndex;

    const sheet =
      model.sheets.find(
        item =>
          item.index ===
          questionsSheetIndex
      );

    if (!sheet) {
      return [];
    }

    const questions = [];

    let startRow = null;
    let number = null;

    for (
      let rowIndex = 0;
      rowIndex <
        sheet.rows.length;
      rowIndex++
    ) {
      const row =
        sheet.rows[
          rowIndex
        ];

      if (
        !this.isQuestionStart(
          row
        )
      ) {
        continue;
      }

      if (
        startRow !== null
      ) {
        questions.push(
          this.createQuestionBlock(
            sheet,
            questions.length,
            number,
            startRow,
            rowIndex - 1
          )
        );
      }

      startRow =
        rowIndex;

      number =
        Number(
          row[0]
        );
    }

    if (
      startRow !== null
    ) {
      questions.push(
        this.createQuestionBlock(
          sheet,
          questions.length,
          number,
          startRow,
          sheet.rows.length - 1
        )
      );
    }

    return questions;
  }

  static createQuestionBlock(
    sheet,
    index,
    number,
    startRow,
    endRow
  ) {
    return {
      index,
      number,
      startRow,
      endRow,

      rows:
        sheet.rows.slice(
          startRow,
          endRow + 1
        )
    };
  }

  static isQuestionStart(
    row
  ) {
    if (
      !row ||
      row.length === 0
    ) {
      return false;
    }

    const questionNumber =
      String(
        row[0]
      ).trim();

    return /^\d+$/.test(
      questionNumber
    );
  }

}


/**
 * ==========================================
 * КОНЕЦ РАЗДЕЛА:
 * ОПРЕДЕЛЕНИЕ БЛОКОВ ВОПРОСОВ
 * ==========================================
 */


/**
 * ==========================================
 * 4. ПЕРЕХОДЫ МЕЖДУ ВОПРОСАМИ
 * ==========================================
 *
 * TransitionParser разбирает значения
 * из Excel.
 *
 * TransitionJsonBuilder преобразует
 * внутренний переход в идентификатор,
 * используемый в итоговом JSON.
 */


class TransitionParser {

  static parse(
    rawTransition
  ) {
    const value =
      String(
        rawTransition || ''
      )
        .trim()
        .toLowerCase();

    if (!value) {
      return null;
    }

    if (
      /^\d+$/.test(
        value
      )
    ) {
      return {
        type: 'question',
        targetId:
          `q_${value}`
      };
    }

    if (
      value === 'конец'
    ) {
      return {
        type: 'end'
      };
    }

    return {
      type: 'unknown',
      rawValue:
        rawTransition
    };
  }

}


class TransitionJsonBuilder {

  static buildNextId(
    transition
  ) {
    if (!transition) {
      return null;
    }

    if (
      transition.type ===
        'question'
    ) {
      return transition.targetId;
    }

    if (
      transition.type ===
        'end'
    ) {
      return 'r_1';
    }

    return null;
  }

}


/**
 * ==========================================
 * КОНЕЦ РАЗДЕЛА:
 * ПЕРЕХОДЫ МЕЖДУ ВОПРОСАМИ
 * ==========================================
 */


/**
 * ==========================================
 * 5. ПАРСЕРЫ ВОПРОСОВ
 * ==========================================
 *
 * QuestionParser определяет тип вопроса.
 *
 * ChoiceQuestionParser содержит общую
 * логику для radio, checkbox и select.
 *
 * Остальные классы разбирают конкретные
 * типы вопросов.
 */


/**
 * Главный маршрутизатор парсеров вопросов.
 */
class QuestionParser {

  static parse(
    questionBlock
  ) {
    this.validateBlock(
      questionBlock
    );

    const type =
      this.detectType(
        questionBlock.rows[0][2]
      );

    let question;

    switch (type) {
      case 'radio':
        question =
          RadioParser.parse(
            questionBlock
          );
        break;

      case 'checkbox':
        question =
          CheckboxParser.parse(
            questionBlock
          );
        break;

      case 'number':
        question =
          NumberParser.parse(
            questionBlock
          );
        break;

      case 'textArea':
        question =
          TextParser.parse(
            questionBlock,
            'textArea'
          );
        break;

      case 'select':
        question =
          SelectParser.parse(
            questionBlock
          );
        break;

      default:
        throw new Error(
          `Неизвестный тип вопроса №${questionBlock.number}.`
        );
    }

    question.description =
      this.extractDescription(
        questionBlock
      );

    return question;
  }

  static extractDescription(
    questionBlock
  ) {
    const descriptions =
      questionBlock.rows
        .slice(1)
        .map(
          row =>
            String(
              row[1] ?? ''
            ).trim()
        )
        .filter(
          Boolean
        );

    return descriptions.join(
      '\n'
    );
  }

  static detectType(
    value
  ) {
    const rawType =
      String(
        value ?? ''
      )
        .replace(
          /[\uFE0E\uFE0F]/g,
          ''
        )
        .trim()
        .toLowerCase();

    if (
      [
        '🔘',
        '◉',
        '○',
        '●',
        '◯',
        '⚪',
        '⦿'
      ].includes(
        rawType
      )
    ) {
      return 'radio';
    }

    if (
      [
        '☑',
        '☐',
        '✓',
        '✔',
        '✅'
      ].includes(
        rawType
      )
    ) {
      return 'checkbox';
    }

    if (
      rawType === 'число'
    ) {
      return 'number';
    }

    if (
      rawType === 'текст' ||
      rawType ===
        'свободная форма'
    ) {
      return 'textArea';
    }

    if (
      [
        '🔽',
        'выпадающий список',
        'select'
      ].includes(
        rawType
      )
    ) {
      return 'select';
    }

    return 'unknown';
  }

  static validateBlock(
    questionBlock
  ) {
    if (
      !questionBlock ||
      !Array.isArray(
        questionBlock.rows
      ) ||
      questionBlock.rows
        .length === 0
    ) {
      throw new Error(
        'Получен некорректный блок вопроса.'
      );
    }

    if (
      !Number.isInteger(
        questionBlock.number
      )
    ) {
      throw new Error(
        'У вопроса отсутствует корректный номер.'
      );
    }
  }

}


/**
 * Общий парсер вопросов с вариантами ответа.
 *
 * Используется для:
 * — radio;
 * — checkbox;
 * — select.
 */
class ChoiceQuestionParser {

  static parse(
    questionBlock,
    type
  ) {
    const question =
      new Question();

    const header =
      questionBlock.rows[0];

    const rawQuestionTransition =
      this.getText(
        header[4]
      );

    question.number =
      questionBlock.number;

    question.id =
      `q_${questionBlock.number}`;

    question.title =
      this.getText(
        header[1]
      );

    question.type =
      type;

    /*
     * Переход из первой строки сохраняем
     * на уровне всего вопроса.
     *
     * Это важно для составных переходов
     * checkbox, записанных в объединённой
     * ячейке Excel.
     */
    question.transition =
      TransitionParser.parse(
        rawQuestionTransition
      );

    /*
     * Проверяем, является ли переход
     * составным checkbox-переходом.
     *
     * Например:
     *
     * Если ТОЛЬКО 7 → №4
     * Если ТОЛЬКО 8 → конец
     * Если ТОЛЬКО (7+8) → №4
     * Во всех остальных случаях → №3
     */
    const compositeTransition =
      type === 'checkbox'
        ? CheckboxParser
            .parseCompositeTransition(
              rawQuestionTransition
            )
        : null;

    question.options =
      questionBlock.rows
        .map(
          row =>
            this.parseOption(
              row,
              {
                rawQuestionTransition,
                hasCompositeTransition:
                  Boolean(
                    compositeTransition
                  )
              }
            )
        )
        .filter(
          Boolean
        );

    return question;
  }

  static parseOption(
    row,
    context
  ) {
    const text =
      this.getText(
        row[3]
      );

    if (!text) {
      return null;
    }

    const option =
      new Option();

    option.text =
      text;

    const rawOptionTransition =
      this.getText(
        row[4]
      );

    const rawQuestionTransition =
      context &&
      typeof context === 'object'
        ? this.getText(
            context
              .rawQuestionTransition
          )
        : '';

    const hasCompositeTransition =
      Boolean(
        context &&
        context
          .hasCompositeTransition
      );

    /*
     * SheetJS заполняет каждую строку
     * объединённой ячейки одним и тем же
     * значением.
     *
     * Если это корректный составной переход
     * checkbox и его текст совпадает с
     * переходом вопроса, не записываем его
     * отдельно в каждый вариант ответа.
     *
     * Сам составной переход уже хранится
     * в question.transition.
     */
    const isRepeatedCompositeTransition =
      hasCompositeTransition &&
      rawOptionTransition &&
      rawOptionTransition ===
        rawQuestionTransition;

    option.transition =
      isRepeatedCompositeTransition
        ? null
        : TransitionParser.parse(
            rawOptionTransition
          );

    option.note =
      this.getText(
        row[5]
      );

    return option;
  }

  static getText(
    value
  ) {
    return String(
      value ?? ''
    ).trim();
  }

}


/**
 * Парсер вопросов radio.
 */
class RadioParser {

  static parse(
    questionBlock
  ) {
    return ChoiceQuestionParser.parse(
      questionBlock,
      'radio'
    );
  }

}


/**
 * Парсер вопросов checkbox
 * и составных checkbox-переходов.
 */
class CheckboxParser {

  static parse(
    questionBlock
  ) {
    return ChoiceQuestionParser.parse(
      questionBlock,
      'checkbox'
    );
  }

  /**
   * Определяет простое условие,
   * записанное в примечании варианта.
   *
   * Поддерживаемые формулировки:
   * — если выбрана только эта галочка;
   * — если выбрана эта галочка;
   * — если выбран только этот вариант;
   * — если выбран этот вариант.
   */
  static parseCondition(
    note
  ) {
    const value =
      String(
        note || ''
      )
        .trim()
        .toLowerCase()
        .replace(
          /ё/g,
          'е'
        );

    if (!value) {
      return null;
    }

    const isOnlyCondition =
      /если выбрана только эта (галочка|опция)/i
        .test(
          value
        ) ||
      /если выбран только этот вариант/i
        .test(
          value
        );

    if (
      isOnlyCondition
    ) {
      return {
        type: 'only'
      };
    }

    const isIncludesCondition =
      /если выбрана эта (галочка|опция)/i
        .test(
          value
        ) ||
      /если выбран этот вариант/i
        .test(
          value
        );

    if (
      isIncludesCondition
    ) {
      return {
        type: 'includes'
      };
    }

    return null;
  }

  /**
   * Проверяет, есть ли у checkbox-вопроса
   * переход, зависящий от примечания
   * конкретного варианта ответа.
   */
  static hasConditionalTransition(
    question
  ) {
    if (
      !question ||
      question.type !==
        'checkbox'
    ) {
      return false;
    }

    const options =
      Array.isArray(
        question.options
      )
        ? question.options
        : [];

    return options.some(
      option =>
        Boolean(
          this.parseCondition(
            option.note
          )
        )
    );
  }

  /**
   * Разбирает составной переход,
   * записанный несколькими строками.
   *
   * Пример:
   *
   * Если ТОЛЬКО 7 → №4
   * Если ТОЛЬКО 8 → конец
   * Если ТОЛЬКО (7+8) → №4
   * Во всех остальных случаях → №3
   */
  static parseCompositeTransition(
    rawTransition
  ) {
    const value =
      String(
        rawTransition || ''
      )
        .trim()
        .toLowerCase()
        .replace(
          /ё/g,
          'е'
        )
        .replace(
          /\r\n?/g,
          '\n'
        );

    if (!value) {
      return null;
    }

    /*
     * Иногда Excel или SheetJS возвращает
     * многострочный текст одной строкой.
     *
     * Добавляем переносы перед ключевыми
     * конструкциями, чтобы парсер работал
     * и с таким форматом.
     */
    const normalizedValue =
      value
        .replace(
          /\s+(?=если\s+только\b)/gi,
          '\n'
        )
        .replace(
          /\s+(?=во всех остальных случаях\b)/gi,
          '\n'
        );

    const lines =
      normalizedValue
        .split(
          '\n'
        )
        .map(
          line =>
            line.trim()
        )
        .filter(
          Boolean
        );

    const branches = [];

    let defaultTargetId =
      null;

    lines.forEach(
      line => {
        const defaultMatch =
          line.match(
            /во всех остальных случаях\s*(?:→|->)\s*(.+)$/i
          );

        if (
          defaultMatch
        ) {
          defaultTargetId =
            this.parseCompositeTarget(
              defaultMatch[1]
            );

          return;
        }

        const conditionMatch =
          line.match(
            /если\s+только\s+\(?([\d+\s]+)\)?\s*(?:→|->)\s*(.+)$/i
          );

        if (
          !conditionMatch
        ) {
          return;
        }

        const optionNumbers =
          conditionMatch[1]
            .split(
              '+'
            )
            .map(
              optionNumber =>
                Number(
                  optionNumber.trim()
                )
            )
            .filter(
              Number.isInteger
            );

        const targetId =
          this.parseCompositeTarget(
            conditionMatch[2]
          );

        if (
          optionNumbers.length === 0 ||
          !targetId
        ) {
          return;
        }

        branches.push({
          /*
           * В Excel варианты нумеруются
           * с единицы, а в JSON значения
           * checkbox хранятся с нуля.
           */
          optionIndexes:
            optionNumbers.map(
              number =>
                number - 1
            ),

          targetId
        });
      }
    );

    if (
      branches.length === 0 ||
      !defaultTargetId
    ) {
      return null;
    }

    return {
      branches,
      defaultTargetId
    };
  }

  /**
   * Преобразует цель составного перехода:
   *
   * №4    → q_4
   * 4     → q_4
   * конец → r_1
   */
  static parseCompositeTarget(
    value
  ) {
    const normalized =
      String(
        value || ''
      )
        .trim()
        .toLowerCase()
        .replace(
          /ё/g,
          'е'
        );

    if (
      normalized === 'конец'
    ) {
      return 'r_1';
    }

    const match =
      normalized.match(
        /№?\s*(\d+)/
      );

    if (!match) {
      return null;
    }

    return `q_${match[1]}`;
  }

}


/**
 * Парсер вопросов select.
 */
class SelectParser {

  static parse(
    questionBlock
  ) {
    return ChoiceQuestionParser.parse(
      questionBlock,
      'select'
    );
  }

}


/**
 * Парсер числовых вопросов.
 */
class NumberParser {

  static parse(
    questionBlock
  ) {
    const question =
      new Question();

    const header =
      questionBlock.rows[0];

    question.number =
      questionBlock.number;

    question.id =
      `q_${questionBlock.number}`;

    question.title =
      String(
        header[1] ?? ''
      ).trim();

    question.type =
      'number';

    question.transition =
      TransitionParser.parse(
        header[4]
      );

    question.note =
      String(
        header[5] ?? ''
      ).trim();

    return question;
  }

}


/**
 * Парсер вопросов со свободным вводом.
 *
 * Используется для:
 * — textInput;
 * — textArea.
 */
class TextParser {

  static parse(
    questionBlock,
    type
  ) {
    const question =
      new Question();

    const header =
      questionBlock.rows[0];

    question.number =
      questionBlock.number;

    question.id =
      `q_${questionBlock.number}`;

    question.title =
      String(
        header[1] ?? ''
      ).trim();

    question.type =
      type;

    question.transition =
      TransitionParser.parse(
        header[4]
      );

    question.note =
      String(
        header[5] ?? ''
      ).trim();

    return question;
  }

}

/**
 * ==========================================
 * 6. ВАЛИДАЦИЯ ВОПРОСОВ
 * ==========================================
 *
 * Проверяет базовую корректность
 * разобранных вопросов.
 */

class QuestionValidator {

  static validate(
    question
  ) {
    const errors = [];

    if (!question.title) {
      errors.push(
        'Не заполнен текст вопроса.'
      );
    }

    if (
      question.type ===
        'unknown'
    ) {
      errors.push(
        'Неизвестный тип вопроса.'
      );
    }

    if (
      [
        'radio',
        'checkbox',
        'select'
      ].includes(
        question.type
      ) &&
      (
        !Array.isArray(
          question.options
        ) ||
        question.options.length === 0
      )
    ) {
      errors.push(
        'Отсутствуют варианты ответа.'
      );
    }

    return errors;
  }

}


/**
 * ==========================================
 * КОНЕЦ РАЗДЕЛА:
 * ВАЛИДАЦИЯ ВОПРОСОВ
 * ==========================================
 */


/**
 * ==========================================
 * 7. ГРАФ И РАСЧЕТ МАРШРУТОВ
 * ==========================================
 *
 * Строит граф переходов, проверяет циклы
 * и рассчитывает длины маршрутов,
 * blockWeight и progressWeight.
 */


class SurveyFlowCalculator {

  constructor(model) {
    this.model = model;
    this.graph = {};
    this.longestPathCache = {};
  }

  buildGraph() {
    this.graph = {};
    this.longestPathCache = {};

    this.model.questions.forEach(
      (
        question,
        questionIndex
      ) => {
        const transitions = [];

        const addedTargets =
          new Set();

        const options =
          Array.isArray(
            question.options
          )
            ? question.options
            : [];

        options.forEach(
          option => {
            const transition =
              option.transition;

            if (!transition) {
              return;
            }

            const targetId =
              this.getTargetId(
                transition
              );

            if (
              !targetId ||
              addedTargets.has(
                targetId
              )
            ) {
              return;
            }

            transitions.push({
              to: targetId,
              transition
            });

            addedTargets.add(
              targetId
            );
          }
        );

        if (
          transitions.length === 0 &&
          question.transition
        ) {
          const targetId =
            this.getTargetId(
              question.transition
            );

          if (targetId) {
            transitions.push({
              to: targetId,

              transition:
                question.transition
            });

            addedTargets.add(
              targetId
            );
          }
        }

        const compositeTransition =
          question.type ===
            'checkbox'
            ? CheckboxParser
                .parseCompositeTransition(
                  question
                    .transition
                    ?.rawValue
                )
            : null;

        if (
          compositeTransition
        ) {
          const compositeTargets = [
            ...compositeTransition
              .branches
              .map(
                branch =>
                  branch.targetId
              ),

            compositeTransition
              .defaultTargetId
          ];

          compositeTargets.forEach(
            targetId => {
              if (
                !targetId ||
                addedTargets.has(
                  targetId
                )
              ) {
                return;
              }

              transitions.push({
                to: targetId,

                transition: {
                  type:
                    'composite',

                  targetId
                }
              });

              addedTargets.add(
                targetId
              );
            }
          );

          question.compositeTransition =
            compositeTransition;
        }

        if (
          transitions.length === 0
        ) {
          const defaultTargetId =
            this.getDefaultTargetId(
              questionIndex
            );

          transitions.push({
            to:
              defaultTargetId,

            transition: {
              type: 'default',

              targetId:
                defaultTargetId
            }
          });
        }

        this.graph[
          question.id
        ] = transitions;
      }
    );

    return this.graph;
  }

  getTargetId(
    transition
  ) {
    if (
      transition.type ===
        'question' ||
      transition.type ===
        'composite' ||
      transition.type ===
        'default'
    ) {
      return transition.targetId;
    }

    if (
      transition.type ===
        'end'
    ) {
      return 'r_1';
    }

    return null;
  }

  getDefaultTargetId(
    questionIndex
  ) {
    const nextQuestion =
      this.model.questions[
        questionIndex + 1
      ];

    if (nextQuestion) {
      return nextQuestion.id;
    }

    return 'r_1';
  }

  getLongestSurveyPath() {
    const firstQuestion =
      this.model.questions[0];

    if (!firstQuestion) {
      return 0;
    }

    return this.getLongestPathFrom(
      firstQuestion.id
    );
  }

  getLongestPathFrom(
    questionId,
    visiting
  ) {
    if (
      questionId === 'r_1'
    ) {
      return 0;
    }

    if (
      Object.prototype
        .hasOwnProperty
        .call(
          this.longestPathCache,
          questionId
        )
    ) {
      return this.longestPathCache[
        questionId
      ];
    }

    const currentVisiting =
      visiting ||
      new Set();

    if (
      currentVisiting.has(
        questionId
      )
    ) {
      throw new Error(
        `Обнаружен циклический переход в вопросе ${questionId}.`
      );
    }

    const transitions =
      this.graph[
        questionId
      ];

    if (!transitions) {
      throw new Error(
        `В графе отсутствует вопрос ${questionId}.`
      );
    }

    currentVisiting.add(
      questionId
    );

    let longestChildPath = 0;

    transitions.forEach(
      transition => {
        const childPath =
          this.getLongestPathFrom(
            transition.to,
            currentVisiting
          );

        longestChildPath =
          Math.max(
            longestChildPath,
            childPath
          );
      }
    );

    currentVisiting.delete(
      questionId
    );

    const pathLength =
      1 + longestChildPath;

    this.longestPathCache[
      questionId
    ] = pathLength;

    return pathLength;
  }

  getRemainingQuestions(
    targetId
  ) {
    return this.getLongestPathFrom(
      targetId
    );
  }

  getProgressWeight(
    questionId,
    targetId
  ) {
    const currentPathLength =
      this.getLongestPathFrom(
        questionId
      );

    const targetPathLength =
      this.getLongestPathFrom(
        targetId
      );

    return (
      currentPathLength -
      targetPathLength
    );
  }

}


/**
 * ==========================================
 * КОНЕЦ РАЗДЕЛА:
 * ГРАФ И РАСЧЕТ МАРШРУТОВ
 * ==========================================
 */


/**
 * ==========================================
 * 8. СБОРКА МОДЕЛИ ОПРОСА
 * ==========================================
 *
 * Центральный класс серверной логики.
 *
 * Собирает вопросы, строит граф,
 * рассчитывает веса и выполняет проверки.
 */


class SurveyBuilder {

  static build(workbookData) {
    const model =
      WorkbookReader.read(
        workbookData
      );

    model.errors =
      Array.isArray(model.errors)
        ? model.errors
        : [];

    model.warnings =
      Array.isArray(model.warnings)
        ? model.warnings
        : [];

    model.questions =
      Array.isArray(model.questions)
        ? model.questions
        : [];

    model.questionBlocks =
      Array.isArray(
        model.questionBlocks
      )
        ? model.questionBlocks
        : [];

    const sheets =
      Array.isArray(model.sheets)
        ? model.sheets
        : [];

    const questionsSheets =
      sheets.filter(
        sheet =>
          sheet.type === 'questions'
      );

    const questionsSheet =
      questionsSheets.find(
        sheet =>
          !sheet.isHidden
      );

    if (!questionsSheet) {
      model.errors.push(
        'Не найден видимый лист «Вопросы».'
      );

      return model;
    }

    model.metadata =
      model.metadata &&
      typeof model.metadata === 'object'
        ? model.metadata
        : {};

    model.metadata.questionsSheetName =
      questionsSheet.name;

    model.metadata.questionsSheetIndex =
      questionsSheet.index;

    model.questionBlocks =
      QuestionDetector.detect(
        model
      );

    model.questions =
      this.buildQuestions(
        model.questionBlocks,
        model
      );

    this.validateUnknownTransitions(
      model
    );

    const flowCalculator =
      new SurveyFlowCalculator(
        model
      );

    model.flowGraph =
      flowCalculator.buildGraph();

    model.metadata.blockWeight =
      flowCalculator
        .getLongestSurveyPath();

    const firstQuestion =
      model.questions[0];

    if (firstQuestion) {
      firstQuestion.blockWeight =
        model.metadata.blockWeight;
    }

    model.questions.forEach(
      question => {
        question.options =
          Array.isArray(
            question.options
          )
            ? question.options
            : [];

        question.remainingQuestions =
          flowCalculator
            .getLongestPathFrom(
              question.id
            );

        const transitions =
          model.flowGraph[
            question.id
          ] || [];

        if (
          transitions.length === 1
        ) {
          question.defaultNextId =
            transitions[0].to;
        }

        question.options.forEach(
          option => {
            const transition =
              option.transition;

            if (!transition) {
              return;
            }

            const targetId =
              flowCalculator
                .getTargetId(
                  transition
                );

            if (!targetId) {
              return;
            }

            option.progressWeight =
              flowCalculator
                .getProgressWeight(
                  question.id,
                  targetId
                );
          }
        );

        if (
          question.compositeTransition &&
          Array.isArray(
            question
              .compositeTransition
              .branches
          )
        ) {
          question
            .compositeTransition
            .branches
            .forEach(
              branch => {
                branch.progressWeight =
                  flowCalculator
                    .getProgressWeight(
                      question.id,
                      branch.targetId
                    );
              }
            );

          const defaultTargetId =
            question
              .compositeTransition
              .defaultTargetId;

          if (defaultTargetId) {
            question
              .compositeTransition
              .defaultProgressWeight =
                flowCalculator
                  .getProgressWeight(
                    question.id,
                    defaultTargetId
                  );
          }
        }
      }
    );

    this
      .deferProgressWeightToConditionalCheckboxes(
        model,
        flowCalculator
      );

    this
      .deferProgressWeightToTargetQuestions(
        model,
        flowCalculator
      );

    /*
     * Проверяем предупреждения только
     * после полного расчёта графа,
     * переходов и весов.
     */
    this.validateFlowWarnings(
      model
    );

    return model;
  }

  static validateUnknownTransitions(
    model
  ) {
    const questions =
      Array.isArray(
        model.questions
      )
        ? model.questions
        : [];

    const warnings =
      new Set();

    questions.forEach(
      question => {
        const questionId =
          question.id ||
          `q_${question.number}`;

        const compositeTransition =
          question.type === 'checkbox'
            ? CheckboxParser
                .parseCompositeTransition(
                  question
                    .transition
                    ?.rawValue
                )
            : null;

        const hasValidCompositeTransition =
          Boolean(
            compositeTransition
          );

        const options =
          Array.isArray(
            question.options
          )
            ? question.options
            : [];

        options.forEach(
          (
            option,
            optionIndex
          ) => {
            const transition =
              option.transition;

            if (
              !transition ||
              transition.type !==
                'unknown' ||
              hasValidCompositeTransition
            ) {
              return;
            }

            const rawValue =
              String(
                transition.rawValue ?? ''
              ).trim();

            warnings.add(
              `В вопросе ${questionId}, ` +
              `вариант ответа ${optionIndex + 1}: ` +
              `не удалось определить переход` +
              (
                rawValue
                  ? ` «${rawValue}».`
                  : '.'
              ) +
              ` Конфиг требует ручной проверки.`
            );
          }
        );

        const questionTransition =
          question.transition;

        if (
          questionTransition &&
          questionTransition.type ===
            'unknown' &&
          !hasValidCompositeTransition
        ) {
          const rawValue =
            String(
              questionTransition
                .rawValue ?? ''
            ).trim();

          warnings.add(
            `В вопросе ${questionId} ` +
            `не удалось определить переход` +
            (
              rawValue
                ? ` «${rawValue}».`
                : '.'
            ) +
            ` Конфиг требует ручной проверки.`
          );
        }
      }
    );

    model.warnings.push(
      ...warnings
    );
  }


  static buildQuestions(
    questionBlocks,
    model
  ) {
    const blocks =
      Array.isArray(
        questionBlocks
      )
        ? questionBlocks
        : [];

    return blocks.map(
      block => {
        const question =
          QuestionParser.parse(
            block
          );

        question.options =
          Array.isArray(
            question.options
          )
            ? question.options
            : [];

        question.errors =
          QuestionValidator
            .validate(
              question
            );

        question.errors =
          Array.isArray(
            question.errors
          )
            ? question.errors
            : [];

        model.errors.push(
          ...question.errors.map(
            error =>
              `Вопрос №${question.number}: ` +
              error
          )
        );

        return question;
      }
    );
  }


  static validateFlowWarnings(
    model
  ) {
    if (
      !Array.isArray(
        model.warnings
      )
    ) {
      model.warnings = [];
    }

    const questions =
      Array.isArray(
        model.questions
      )
        ? model.questions
        : [];

    const graph =
      model.flowGraph || {};

    const questionIndexes = {};
    const incomingSources = {};

    questions.forEach(
      (question, index) => {
        questionIndexes[
          question.id
        ] = index;

        incomingSources[
          question.id
        ] = [];
      }
    );

    const warnings =
      new Set(
        model.warnings
      );

    Object.entries(
      graph
    ).forEach(
      ([
        sourceId,
        transitions
      ]) => {
        const sourceIndex =
          questionIndexes[
            sourceId
          ];

        if (
          sourceIndex === undefined ||
          !Array.isArray(
            transitions
          )
        ) {
          return;
        }

        transitions.forEach(
          transition => {
            const targetId =
              transition.to;

            if (!targetId) {
              warnings.add(
                `У вопроса ${sourceId} ` +
                `есть переход без цели.`
              );

              return;
            }

            if (
              targetId === 'r_1'
            ) {
              return;
            }

            const targetIndex =
              questionIndexes[
                targetId
              ];

            if (
              targetIndex === undefined
            ) {
              warnings.add(
                `Вопрос ${sourceId} ведет ` +
                `к отсутствующему вопросу ` +
                `${targetId}.`
              );

              return;
            }

            incomingSources[
              targetId
            ].push(
              sourceId
            );

            if (
              targetIndex <
              sourceIndex
            ) {
              warnings.add(
                `Вопрос ${sourceId} ведет назад ` +
                `к более раннему вопросу ` +
                `${targetId}. Проверьте расчет ` +
                `blockWeight и ` +
                `progressWeightPrecalculate.`
              );
            }
          }
        );
      }
    );

    /*
     * Проверяем вопросы, в которые
     * можно попасть как из более ранней,
     * так и из более поздней части опроса.
     */
    Object.entries(
      incomingSources
    ).forEach(
      ([
        targetId,
        sourceIds
      ]) => {
        const targetIndex =
          questionIndexes[
            targetId
          ];

        const uniqueSourceIds = [
          ...new Set(
            sourceIds
          )
        ];

        const hasEarlierSource =
          uniqueSourceIds.some(
            sourceId =>
              questionIndexes[
                sourceId
              ] < targetIndex
          );

        const hasLaterSource =
          uniqueSourceIds.some(
            sourceId =>
              questionIndexes[
                sourceId
              ] > targetIndex
          );

        if (
          hasEarlierSource &&
          hasLaterSource
        ) {
          warnings.add(
            `Вопрос ${targetId} имеет ` +
            `несколько входящих переходов: ` +
            `${uniqueSourceIds.join(', ')}. ` +
            `Автоматический расчет ` +
            `progressWeight может быть ` +
            `недостаточен. Проверьте ` +
            `progressWeightPrecalculate ` +
            `вручную.`
          );
        }
      }
    );

    /*
     * Первый вопрос не обязан иметь
     * входящий переход. Остальные должны
     * быть доступны хотя бы из одного
     * вопроса.
     */
    questions.forEach(
      (question, index) => {
        if (index === 0) {
          return;
        }

        const sources =
          incomingSources[
            question.id
          ] || [];

        if (
          sources.length === 0
        ) {
          warnings.add(
            `Вопрос ${question.id} ` +
            `недостижим: к нему не ведет ` +
            `ни один переход.`
          );
        }
      }
    );

    /*
     * Проверяем итоговые вопросы,
     * у которых не обнаружено переходов.
     */
    questions.forEach(
      question => {
        const transitions =
          graph[
            question.id
          ];

        if (
          Array.isArray(
            transitions
          ) &&
          transitions.length > 0
        ) {
          return;
        }

        const hasDefaultNextId =
          Boolean(
            question.defaultNextId
          );

        const hasPrecalculatedNextId =
          Boolean(
            question
              .nextIdPrecalculate
          );

        const hasCompositeTransition =
          Boolean(
            question
              .compositeTransition
          );

        if (
          !hasDefaultNextId &&
          !hasPrecalculatedNextId &&
          !hasCompositeTransition
        ) {
          warnings.add(
            `У вопроса ${question.id} ` +
            `не определен переход к ` +
            `следующему вопросу или результату.`
          );
        }
      }
    );

    model.warnings = [
      ...warnings
    ];
  }


  static deferProgressWeightToConditionalCheckboxes(
    model,
    flowCalculator
  ) {
    const questionsById = {};

    model.questions.forEach(
      question => {
        questionsById[
          question.id
        ] = question;
      }
    );

    model.questions.forEach(
      sourceQuestion => {
        const options =
          Array.isArray(
            sourceQuestion.options
          )
            ? sourceQuestion.options
            : [];

        options.forEach(
          option => {
            const transition =
              option.transition;

            if (!transition) {
              return;
            }

            const targetId =
              flowCalculator
                .getTargetId(
                  transition
                );

            if (!targetId) {
              return;
            }

            const targetQuestion =
              questionsById[
                targetId
              ];

            if (
              !targetQuestion ||
              !CheckboxParser
                .hasConditionalTransition(
                  targetQuestion
                )
            ) {
              return;
            }

            const incomingProgressWeight =
              option.progressWeight || 1;

            const progressWeightOffset =
              Math.max(
                0,
                incomingProgressWeight - 1
              );

            if (
              progressWeightOffset === 0
            ) {
              return;
            }

            const existingOffset =
              targetQuestion
                .progressWeightOffset || 0;

            if (
              existingOffset !== 0 &&
              existingOffset !==
                progressWeightOffset
            ) {
              throw new Error(
                `У вопроса ${targetId} ` +
                `обнаружены входящие ` +
                `переходы с разным весом ` +
                `прогресса.`
              );
            }

            targetQuestion
              .progressWeightOffset =
                progressWeightOffset;

            option.progressWeight = 1;
          }
        );
      }
    );
  }


  static deferProgressWeightToTargetQuestions(
    model,
    flowCalculator
  ) {
    const questionsById = {};
    const incomingSourcesByTarget =
      {};

    model.questions.forEach(
      question => {
        questionsById[
          question.id
        ] = question;
      }
    );

    Object.keys(
      model.flowGraph || {}
    ).forEach(
      sourceId => {
        const transitions =
          model.flowGraph[
            sourceId
          ] || [];

        transitions.forEach(
          transition => {
            const targetId =
              transition.to;

            if (
              !targetId ||
              targetId === 'r_1'
            ) {
              return;
            }

            if (
              !incomingSourcesByTarget[
                targetId
              ]
            ) {
              incomingSourcesByTarget[
                targetId
              ] = new Set();
            }

            incomingSourcesByTarget[
              targetId
            ].add(
              sourceId
            );
          }
        );
      }
    );

    model.questions.forEach(
      sourceQuestion => {
        const optionsByTarget = {};

        const sourceOptions =
          Array.isArray(
            sourceQuestion.options
          )
            ? sourceQuestion.options
            : [];

        sourceOptions.forEach(
          option => {
            const transition =
              option.transition;

            if (!transition) {
              return;
            }

            const targetId =
              flowCalculator
                .getTargetId(
                  transition
                );

            if (
              !targetId ||
              targetId === 'r_1'
            ) {
              return;
            }

            if (
              !optionsByTarget[
                targetId
              ]
            ) {
              optionsByTarget[
                targetId
              ] = [];
            }

            optionsByTarget[
              targetId
            ].push(
              option
            );
          }
        );

        Object.keys(
          optionsByTarget
        ).forEach(
          targetId => {
            const targetQuestion =
              questionsById[
                targetId
              ];

            if (!targetQuestion) {
              return;
            }

            /*
             * Условные чекбоксы
             * обрабатываются отдельным
             * методом.
             */
            if (
              CheckboxParser
                .hasConditionalTransition(
                  targetQuestion
                )
            ) {
              return;
            }

            const incomingSources =
              incomingSourcesByTarget[
                targetId
              ];

            /*
             * Перенос безопасен только
             * при одном источнике перехода.
             */
            if (
              !incomingSources ||
              incomingSources.size !== 1 ||
              !incomingSources.has(
                sourceQuestion.id
              )
            ) {
              return;
            }

            const targetOptions =
              optionsByTarget[
                targetId
              ];

            const progressWeights =
              targetOptions.map(
                option =>
                  option
                    .progressWeight || 1
              );

            const firstProgressWeight =
              progressWeights[0];

            if (
              !firstProgressWeight ||
              firstProgressWeight <= 1
            ) {
              return;
            }

            const hasDifferentWeights =
              progressWeights.some(
                progressWeight =>
                  progressWeight !==
                  firstProgressWeight
              );

            if (
              hasDifferentWeights
            ) {
              throw new Error(
                `Переходы из вопроса ` +
                `${sourceQuestion.id} в ` +
                `${targetId} имеют разный ` +
                `вес прогресса.`
              );
            }

            targetQuestion
              .progressWeight =
                firstProgressWeight;

            targetOptions.forEach(
              option => {
                option.progressWeight =
                  1;
              }
            );
          }
        );
      }
    );
  }

}


/**
 * ==========================================
 * КОНЕЦ РАЗДЕЛА:
 * СБОРКА МОДЕЛИ ОПРОСА
 * ==========================================
 */


/**
 * ==========================================
 * 9. ТИПОГРАФИКА И СБОРКА ВОПРОСОВ JSON
 * ==========================================
 *
 * Форматирует текст и преобразует
 * внутренние модели вопросов в JSON-карточки.
 */


class QuestionJsonBuilder {

  static formatNumericRanges(value) {
    return String(value ?? '')
      .replace(
        /(\d)[ \t\u00A0]*[-—–][ \t\u00A0]*(\d)/g,
        '$1\u2060—\u2060$2'
      );
  }

  static formatMeasurementUnits(value) {
    const units =
      [
        'кВт[·⋅]?ч',
        'кВт',
        'МВт',
        'Вт',

        'км²',
        'дм²',
        'см²',
        'мм²',
        'м²',

        'км³',
        'дм³',
        'см³',
        'мм³',
        'м³',

        'км',
        'дм',
        'см',
        'мм',
        'м',

        'мг',
        'кг',
        'г',
        'т',

        'мкл',
        'мл',
        'дл',
        'л',

        'га'
      ].join('|');

    let result =
      String(value ?? '');

    /*
    * Единицы измерения:
    * 80 м² → 80 м²
    * 15 кг → 15 кг
    * 20 кВт·ч → 20 кВт·ч
    */
    result = result.replace(
      new RegExp(
        `(\\d)[ \\t\\u00A0]*(${units})` +
        `(?=[\\s.,!?;:)]|$)`,
        'gi'
      ),
      '$1\u00A0$2'
    );

    /*
    * Проценты:
    * 15 % → 15 %
    */
    result = result.replace(
      /(\d)[ \t\u00A0]*%/g,
      '$1\u00A0%'
    );

    /*
    * Температура:
    * 20 °C → 20 °C
    * 20 °С → 20 °С
    */
    result = result.replace(
      /(\d)[ \t\u00A0]*°[ \t\u00A0]*([CС])/g,
      '$1\u00A0°$2'
    );

    /*
    * Символы валют:
    * 1000 ₽ → 1000 ₽
    * 500 $ → 500 $
    * 200 € → 200 €
    */
    result = result.replace(
      /(\d)[ \t\u00A0]*(₽|\$|€|£|¥)/g,
      '$1\u00A0$2'
    );

    /*
    * Сокращения валют:
    * 1000 руб. → 1000 руб.
    * 500 долл. → 500 долл.
    * 200 евро → 200 евро
    */
    result = result.replace(
      /(\d)[ \t\u00A0]*(руб\.?|долл\.?|евро)(?=[\s.,!?;:)]|$)/gi,
      '$1\u00A0$2'
    );

    /*
    * Год:
    * 2026 г. → 2026 г.
    *
    * Не затрагивает единицу массы:
    * 500 г
    */
    result = result.replace(
      /\b((?:19|20)\d{2})[ \t\u00A0]*г\.(?=[\s.,!?;:)]|$)/g,
      '$1\u00A0г.'
    );

    /*
    * Год словами:
    * 2026 год → 2026 год
    * 2026 года → 2026 года
    * в 2026 году → в 2026 году
    * за 5 лет → за 5 лет
    */
    result = result.replace(
      /(\d)[ \t\u00A0]+(год|года|году|годом|годы|годах|лет)(?=[\s.,!?;:)]|$)/gi,
      '$1\u00A0$2'
    );

    return result;
  }
  
  static formatShortHyphenatedWords(
    value
  ) {
    const pattern =
      /(^|[^А-ЯЁа-яёA-Za-z0-9-])([А-ЯЁа-яёA-Za-z0-9]+)-([А-ЯЁа-яёA-Za-z0-9]+)([ \t\u00A0]+|(?=$|[^А-ЯЁа-яёA-Za-z0-9-]))/g;

    return String(
      value ?? ''
    ).replace(
      pattern,
      function (
        match,
        prefix,
        firstPart,
        secondPart,
        whitespace
      ) {
        const hasShortPart =
          firstPart.length <= 4 ||
          secondPart.length <= 4;

        if (!hasShortPart) {
          return match;
        }

        const gluedWord =
          firstPart +
          '\u2060-\u2060' +
          secondPart;

        if (
          whitespace &&
          /^[ \t\u00A0]+$/.test(
            whitespace
          )
        ) {
          return (
            prefix +
            gluedWord +
            ' '
          );
        }

        return (
          prefix +
          gluedWord
        );
      }
    );
  }


  static formatText(value) {
    let result =
      String(value ?? '');

    result =
      this.formatNumericRanges(
        result
      );

    result =
      this.formatMeasurementUnits(
        result
      );

    result =
      this.formatShortHyphenatedWords(
        result
      );

    return result
      .replace(
        /Т(?:\u2060)?—(?:\u2060)?Ж/g,
        'Т\u2060—\u2060Ж'
      )
      .replace(
        /([А-ЯЁа-яёA-Za-z]+)[ \t\u00A0]+(же|ли|бы)[ \t\u00A0]+/gi,
        '$1\u00A0$2 \u200B'
      )
      .replace(
        /([А-ЯЁа-яёA-Za-z]+)[ \t\u00A0]+(же|ли|бы)(?=[.,!?;:]|$)/gi,
        '$1\u00A0$2'
      );
  }

  static build(question) {
    let result;

    switch (question.type) {
      case 'radio':
        result =
          this.buildRadio(question);
        break;

      case 'checkbox':
        result =
          this.buildCheckboxes(question);
        break;

      case 'number':
        result =
          this.buildNumberInput(question);
        break;

      case 'textArea':
        result =
          this.buildTextArea(question);
        break;

      case 'select':
        result =
          this.buildSelect(question);
        break;

      default:
        throw new Error(
          `Неизвестный тип вопроса: ${question.type}`
        );
    }

    if (question.progressWeight > 1) {
      result.progressWeight =
        question.progressWeight;
    }

    return result;
  }

  static buildRadio(question) {
    const defaultNextId =
      TransitionJsonBuilder.buildNextId(
        question.options[0]?.transition
      );

    const options =
      question.options.map(option => {
        const result = {
          label: this.formatText(
            option.text
          )
        };

        const optionNextId =
          TransitionJsonBuilder.buildNextId(
            option.transition
          );

        if (
          optionNextId &&
          optionNextId !== defaultNextId
        ) {
          result.defaultNextId =
            optionNextId;
        }

        if (option.progressWeight > 1) {
          result.progressWeight =
            option.progressWeight;
        }

        return result;
      });

    const result = {
      type: 'radio',
      heading: this.formatText(
        question.title
      ),
      options: options,
      isRequired: true
    };

    if (question.blockWeight) {
      result.blockWeight =
        question.blockWeight;
    }

    if (question.description) {
      result.description =
        this.formatText(
          question.description
        );
    }

    result.isCollected = true;
    result.collectedKey =
      `q_${question.number}`;

    if (defaultNextId) {
      result.defaultNextId =
        defaultNextId;
    }

    return result;
  }

  static buildSelect(question) {
    const otherOption =
      question.options.find(
        option =>
          this.isSelectOtherOption(
            option
          )
      );

    const regularOptions =
      question.options.filter(
        option =>
          !this.isSelectOtherOption(
            option
          )
      );

    const defaultNextId =
      TransitionJsonBuilder.buildNextId(
        regularOptions[0]?.transition
      );

    const options =
      regularOptions.map(option => {
        const result = {
          label: this.formatText(
            option.text
          )
        };

        const optionNextId =
          TransitionJsonBuilder.buildNextId(
            option.transition
          );

        /*
        * Сохраняем отдельный переход,
        * только если он отличается от
        * общего перехода вопроса.
        */
        if (
          optionNextId &&
          optionNextId !== defaultNextId
        ) {
          result.defaultNextId =
            optionNextId;
        }

        return result;
      });

    const result = {
      type: 'select',

      heading: this.formatText(
        question.title
      ),

      options,

      isCollected: true,

      collectedKey:
        question.id
    };

    if (
      this.isRequiredQuestion(
        question
      )
    ) {
      result.isRequired = true;
    }

    if (question.description) {
      result.description =
        this.formatText(
          question.description
        );
    }

    if (otherOption) {
      result.otherOption = {
        label: this.formatText(
          otherOption.text
        )
      };

      result.otherMaxLength = 1000;
    }

    if (defaultNextId) {
      result.defaultNextId =
        defaultNextId;
    }

    return result;
  }

  static buildCheckboxes(question) {
    const otherOption =
      question.options.find(
        option =>
          this.isOtherOption(option)
      );

    const regularOptions =
      question.options.filter(
        option =>
          !this.isOtherOption(option)
      );

    const result = {
      type: 'checkboxes',
      heading: this.formatText(
        question.title
      ),
      options:
        regularOptions.map(
          option =>
            this.formatText(
              option.text
            )
        ),
      isRequired: true
    };

    if (question.description) {
      result.description =
        this.formatText(
          question.description
        );
    }

    result.isCollected = true;
    result.collectedKey =
      `q_${question.number}`;

    if (otherOption) {
      result.otherOption =
        this.formatText(
          otherOption.text
        );

      result.otherMaxLength = 1000;
    }

    const defaultNextId =
      this.getCheckboxDefaultNextId(
        regularOptions
      );

    const conditionalTransition =
      this.buildCheckboxCompositeTransition(
        question
      ) ||
      this.buildCheckboxConditionalTransition(
        question,
        regularOptions,
        defaultNextId
      ) ||
      this.buildCheckboxTransitionByTargets(
        question,
        regularOptions,
        defaultNextId
      );

    if (conditionalTransition) {
      result.nextIdPrecalculate =
        conditionalTransition
          .nextIdPrecalculate;

      if (
        conditionalTransition
          .progressWeightPrecalculate
      ) {
        result.progressWeightPrecalculate =
          conditionalTransition
            .progressWeightPrecalculate;
      }
    } else if (defaultNextId) {
      result.defaultNextId =
        defaultNextId;
    }

    return result;
  }

  static isOtherOption(option) {
    const note = String(
      option.note || ''
    )
      .trim()
      .toLowerCase()
      .replace(/ё/g, 'е');

    return (
      note.includes('свой вариант ответа') ||
      note.includes('свободная форма')
    );
  }

  static isAgeQuestion(question) {
    const title =
      String(question.title || '')
        .trim()
        .toLowerCase()
        .replace(/ё/g, 'е');

    return (
      title.includes('возраст') ||
      /сколько\s+(вам|тебе)\s+лет/i.test(
        title
      )
    );
  }

  static getCurrencyConfig(question) {
    const note = String(
      question.note || ''
    )
      .trim()
      .toUpperCase()
      .replace(/Ё/g, 'Е');

    const currencies = {
      RUB: {
        markers: [
          'RUB',
          'РУБЛЬ',
          'РУБЛИ',
          'РУБЛЕЙ'
        ],
        suffix: {
          one: 'рубль',
          two: 'рубля',
          five: 'рублей'
        },
        safeMaxThreshold: 999999
      },

      USD: {
        markers: [
          'USD',
          'ДОЛЛАР',
          'ДОЛЛАРЫ',
          'ДОЛЛАРОВ'
        ],
        suffix: {
          one: 'доллар',
          two: 'доллара',
          five: 'долларов'
        },
        safeMaxThreshold: 999999
      },

      EUR: {
        markers: [
          'EUR',
          'ЕВРО'
        ],
        suffix: {
          one: 'евро',
          two: 'евро',
          five: 'евро'
        },
        safeMaxThreshold: 999999
      }
    };

    return Object.values(
      currencies
    ).find(currency =>
      currency.markers.includes(note)
    ) || null;
  }

  static buildNumberInput(question) {
    const questionId =
      question.id;

    const nextQuestionId =
      question.defaultNextId;

    const result = {
      type: 'numberInput'
    };

    const isAge =
      this.isAgeQuestion(
        question
      );

    const currency =
      this.getCurrencyConfig(
        question
      );

    const numberRange =
      this.getNumberRange(
        question
      );

    if (isAge) {
      result.suffix = {
        one: 'год',
        two: 'года',
        five: 'лет'
      };
    } else if (currency) {
      result.suffix =
        currency.suffix;
    }

    result.heading =
      this.formatText(
        question.title
      );

    result.isRequired = true;

    if (question.description) {
      result.description =
        this.formatText(
          question.description
        );
    }

    result.isCollected = true;

    result.collectedKey =
      questionId;

    if (numberRange) {
      result.safeMaxThreshold =
        numberRange.max;
    } else if (isAge) {
      result.safeMaxThreshold = 99;
    } else if (currency) {
      result.safeMaxThreshold =
        currency.safeMaxThreshold;
    }

    if (nextQuestionId) {
      const condition =
        numberRange
          ? `keys.${questionId} >= ${numberRange.min}`
          : `keys.${questionId} > 0`;

      result.nextIdPrecalculate =
        `${condition} ` +
        `? '${nextQuestionId}' : null`;
    }

    return result;
  }

  static getCheckboxDefaultNextId(
    options
  ) {
    const targetCounts = {};

    options.forEach(option => {
      const targetId =
        TransitionJsonBuilder.buildNextId(
          option.transition
        );

      if (!targetId) {
        return;
      }

      targetCounts[targetId] =
        (targetCounts[targetId] || 0) + 1;
    });

    const targets =
      Object.keys(targetCounts);

    if (targets.length === 0) {
      return null;
    }

    targets.sort(
      (firstTarget, secondTarget) =>
        targetCounts[secondTarget] -
        targetCounts[firstTarget]
    );

    return targets[0];
  }

  static buildCheckboxConditionalTransition(
    question,
    options,
    defaultNextId
  ) {
    const conditionalOptionIndex =
      options.findIndex(option =>
        Boolean(
          CheckboxParser.parseCondition(
            option.note
          )
        )
      );

    if (
      conditionalOptionIndex === -1
    ) {
      return null;
    }

    const conditionalOption =
      options[
        conditionalOptionIndex
      ];

    const checkboxCondition =
      CheckboxParser.parseCondition(
        conditionalOption.note
      );

    const conditionalTargetId =
      TransitionJsonBuilder.buildNextId(
        conditionalOption.transition
      );

    if (
      !checkboxCondition ||
      !conditionalTargetId ||
      !defaultNextId ||
      conditionalTargetId ===
        defaultNextId
    ) {
      return null;
    }

    const defaultOption =
      options.find(
        (
          option,
          optionIndex
        ) => {
          if (
            optionIndex ===
            conditionalOptionIndex
          ) {
            return false;
          }

          const optionTargetId =
            TransitionJsonBuilder.buildNextId(
              option.transition
            );

          return (
            optionTargetId ===
            defaultNextId
          );
        }
      );

    const progressWeightOffset =
      question.progressWeightOffset || 0;

    const conditionalProgressWeight =
      (conditionalOption.progressWeight || 1) +
      progressWeightOffset;

    const defaultProgressWeight =
      (defaultOption?.progressWeight || 1) +
      progressWeightOffset;

    const questionId =
      question.id;

    const includesExpression =
      `keys.${questionId}.includes(` +
      `'${conditionalOptionIndex}')`;

    let condition;

    if (
      checkboxCondition.type ===
      'only'
    ) {
      condition =
        `keys.${questionId}.length === 1 && ` +
        includesExpression;
    } else {
      condition =
        includesExpression;
    }

    const result = {
      nextIdPrecalculate:
        `${condition} ` +
        `? '${conditionalTargetId}' ` +
        `: '${defaultNextId}'`
    };

    if (
      conditionalProgressWeight > 1 ||
      defaultProgressWeight > 1
    ) {
      result.progressWeightPrecalculate =
        `${condition} ` +
        `? ${conditionalProgressWeight} ` +
        `: ${defaultProgressWeight}`;
    }

    return result;
  }



  static buildCheckboxTransitionByTargets(
    question,
    options,
    defaultNextId
  ) {
    if (
      !defaultNextId ||
      !Array.isArray(options) ||
      options.length === 0
    ) {
      return null;
    }

    const optionsByTarget = {};

    options.forEach(
      (
        option,
        optionIndex
      ) => {
        const targetId =
          TransitionJsonBuilder.buildNextId(
            option.transition
          );

        if (!targetId) {
          return;
        }

        if (!optionsByTarget[targetId]) {
          optionsByTarget[targetId] = [];
        }

        optionsByTarget[targetId].push({
          option,
          optionIndex
        });
      }
    );

    const targetIds =
      Object.keys(optionsByTarget);

    /*
    * Обычный последовательный переход:
    * все варианты ведут в одно место.
    */
    if (targetIds.length <= 1) {
      return null;
    }

    /*
    * Пока автоматически поддерживаем
    * только развилку между двумя целями.
    *
    * Если целей три или больше, без явной
    * заметки нельзя надежно определить
    * нужную формулу.
    */
    if (targetIds.length !== 2) {
      return null;
    }

    const conditionalTargetId =
      targetIds.find(
        targetId =>
          targetId !== defaultNextId
      );

    if (!conditionalTargetId) {
      return null;
    }

    const conditionalItems =
      optionsByTarget[
        conditionalTargetId
      ];

    const defaultItems =
      optionsByTarget[
        defaultNextId
      ];

    if (
      !conditionalItems?.length ||
      !defaultItems?.length
    ) {
      return null;
    }

    const questionId =
      question.id;

    const condition =
      conditionalItems
        .map(
          item =>
            `keys.${questionId}.includes(` +
            `'${item.optionIndex}')`
        )
        .join(' || ');

    const progressWeightOffset =
      question.progressWeightOffset || 0;

    const conditionalProgressWeight =
      (
        conditionalItems[0]
          .option
          .progressWeight || 1
      ) +
      progressWeightOffset;

    const defaultProgressWeight =
      (
        defaultItems[0]
          .option
          .progressWeight || 1
      ) +
      progressWeightOffset;

    const result = {
      nextIdPrecalculate:
        `${condition} ` +
        `? '${conditionalTargetId}' ` +
        `: '${defaultNextId}'`
    };

    if (
      conditionalProgressWeight > 1 ||
      defaultProgressWeight > 1
    ) {
      result.progressWeightPrecalculate =
        `${condition} ` +
        `? ${conditionalProgressWeight} ` +
        `: ${defaultProgressWeight}`;
    }

    return result;
  }

  static buildTextArea(question) {
    const result = {
      type: 'textArea',

      heading: this.formatText(
        question.title
      ),

      maxLength: 1000,

      isCollected: true,

      collectedKey:
        question.id
    };

    if (
      this.isRequiredQuestion(
        question
      )
    ) {
      result.isRequired = true;
    }

    if (question.description) {
      result.description =
        this.formatText(
          question.description
        );
    }

    if (question.defaultNextId) {
      result.defaultNextId =
        question.defaultNextId;
    }

    return result;
  }

  static isRequiredQuestion(question) {
    const note = String(
      question.note || ''
    )
      .trim()
      .toLowerCase()
      .replace(/ё/g, 'е');

    const isOptional =
      /не\s*обязательный\s+вопрос/i
        .test(note);

    return !isOptional;
  }

  static getNumberRange(question) {
    const note = String(
      question.note || ''
    )
      .trim()
      .toLowerCase()
      .replace(/ё/g, 'е')
      .replace(/\u00a0/g, ' ');

    const match = note.match(
      /(?:^|\s)от\s+(-?\d+)\s+до\s+(-?\d+)(?:\s|$)/
    );

    if (!match) {
      return null;
    }

    const min = Number(
      match[1]
    );

    const max = Number(
      match[2]
    );

    if (
      !Number.isFinite(min) ||
      !Number.isFinite(max) ||
      min > max
    ) {
      return null;
    }

    return {
      min,
      max
    };
  }

  static buildCheckboxCompositeTransition(
    question
  ) {
    const compactTransition =
      this.buildCompactCompositeTransition(
        question
      );

    if (compactTransition) {
      return compactTransition;
    }

    const composite =
      question.compositeTransition;

    if (
      !composite ||
      !Array.isArray(
        composite.branches
      ) ||
      !composite.defaultTargetId
    ) {
      return null;
    }

    const questionId =
      question.id;

    const groupedBranches = {};

    composite.branches.forEach(
      branch => {
        const key =
          `${branch.targetId}|` +
          `${branch.progressWeight}`;

        if (!groupedBranches[key]) {
          groupedBranches[key] = {
            targetId:
              branch.targetId,
            progressWeight:
              branch.progressWeight,
            conditions: []
          };
        }

        const indexes =
          branch.optionIndexes;

        const conditionParts = [
          `keys.${questionId}.length === ${indexes.length}`
        ];

        indexes.forEach(index => {
          conditionParts.push(
            `keys.${questionId}.includes('${index}')`
          );
        });

        groupedBranches[key]
          .conditions
          .push(
            conditionParts.join(' && ')
          );
      }
    );

    const branches =
      Object.values(
        groupedBranches
      );

    let nextExpression =
      `'${composite.defaultTargetId}'`;

    let progressExpression =
      String(
        composite.defaultProgressWeight
      );

    branches
      .slice()
      .reverse()
      .forEach(branch => {
        const condition =
          branch.conditions
            .map(value =>
              `(${value})`
            )
            .join(' || ');

        nextExpression =
          `${condition} ` +
          `? '${branch.targetId}' ` +
          `: ${nextExpression}`;

        progressExpression =
          `${condition} ` +
          `? ${branch.progressWeight} ` +
          `: ${progressExpression}`;
      });

    return {
      nextIdPrecalculate:
        nextExpression,

      progressWeightPrecalculate:
        progressExpression
    };
  }

  static buildCompactCompositeTransition(
    question
  ) {
    const composite =
      question.compositeTransition;

    if (
      !composite ||
      !Array.isArray(composite.branches) ||
      !Array.isArray(question.options)
    ) {
      return null;
    }

    const branches =
      composite.branches;

    /*
    * Ищем ветку с двумя выбранными
    * вариантами: например [6, 7] → q_4.
    */
    const pairBranch =
      branches.find(
        branch =>
          Array.isArray(
            branch.optionIndexes
          ) &&
          branch.optionIndexes.length === 2
      );

    if (!pairBranch) {
      return null;
    }

    const pairIndexes =
      pairBranch.optionIndexes
        .slice()
        .sort(
          (
            firstIndex,
            secondIndex
          ) =>
            firstIndex - secondIndex
        );

    const firstIndex =
      pairIndexes[0];

    const secondIndex =
      pairIndexes[1];

    /*
    * Компактная формула безопасна,
    * только если особые варианты —
    * последние два в списке.
    */
    if (
      secondIndex !==
        question.options.length - 1 ||
      firstIndex !== secondIndex - 1
    ) {
      return null;
    }

    const firstSingleBranch =
      branches.find(
        branch =>
          branch.optionIndexes.length === 1 &&
          branch.optionIndexes[0] ===
            firstIndex
      );

    const secondSingleBranch =
      branches.find(
        branch =>
          branch.optionIndexes.length === 1 &&
          branch.optionIndexes[0] ===
            secondIndex
      );

    if (
      !firstSingleBranch ||
      !secondSingleBranch
    ) {
      return null;
    }

    /*
    * Выбор обоих вариантов должен вести
    * туда же, куда первый особый вариант.
    */
    if (
      pairBranch.targetId !==
        firstSingleBranch.targetId ||
      pairBranch.progressWeight !==
        firstSingleBranch.progressWeight
    ) {
      return null;
    }

    /*
    * Второй особый вариант должен вести
    * в другую ветку.
    */
    if (
      secondSingleBranch.targetId ===
        firstSingleBranch.targetId
    ) {
      return null;
    }

    const questionId =
      question.id;

    const defaultTargetId =
      composite.defaultTargetId;

    const defaultProgressWeight =
      composite.defaultProgressWeight;

    const firstTargetId =
      firstSingleBranch.targetId;

    const firstProgressWeight =
      firstSingleBranch.progressWeight;

    const secondTargetId =
      secondSingleBranch.targetId;

    const secondProgressWeight =
      secondSingleBranch.progressWeight;

    const defaultCondition =
      `keys.${questionId}.some(` +
      `(key) => key < ${firstIndex})`;

    const firstCondition =
      `keys.${questionId}.includes(` +
      `'${firstIndex}')`;

    return {
      nextIdPrecalculate:
        `${defaultCondition} ` +
        `? '${defaultTargetId}' ` +
        `: (${firstCondition} ` +
        `? '${firstTargetId}' ` +
        `: '${secondTargetId}')`,

      progressWeightPrecalculate:
        `${defaultCondition} ` +
        `? ${defaultProgressWeight} ` +
        `: (${firstCondition} ` +
        `? ${firstProgressWeight} ` +
        `: ${secondProgressWeight})`
    };
  }

  static isSelectOtherOption(
    option
  ) {
    const note = String(
      option.note || ''
    )
      .trim()
      .toLowerCase()
      .replace(/ё/g, 'е');

    return (
      note.includes(
        'свой вариант ответа'
      ) ||
      note.includes(
        'свободная форма'
      ) ||
      note.includes(
        'возможность вписать'
      ) ||
      note.includes(
        'можно вписать'
      )
    );
  }
}

/**
 * ==========================================
 * КОНЕЦ РАЗДЕЛА:
 * ТИПОГРАФИКА И СБОРКА ВОПРОСОВ JSON
 * ==========================================
 */


/**
 * ==========================================
 * 10. СБОРКА ИТОГОВОГО JSON
 * ==========================================
 *
 * Формирует корневую структуру конфига:
 * auth, results, questions, historyKey
 * и остальные обязательные поля.
 */


class JsonBuilder {

  static build(model) {
      const surveySlug =
        model.metadata.surveySlug;

      return {
        auth: {
          popupTitle: '',
          loginBtnLabel:
            'Войдите, чтобы посчитать',
          pointOfContact: surveySlug,
          isLoginRequired: false,
          popupDescription: ''
        },

        uiText: {},

        apiPath: surveySlug,

        results: {
          r_1: {
            heading:
              QuestionJsonBuilder.formatText(
                model.metadata.resultHeading
              ),

            description:
              this.buildResultDescription(
                model
              )
          }
        },

        extraKeys: {},

        questions:
          this.buildQuestions(model),

        historyKey: surveySlug,

        blocksCount: 1,

        genericCards: {},

        initialCardKey: 'q_1'
      };
  }

  static buildResultDescription(
    model
  ) {
    const description =
      QuestionJsonBuilder.formatText(
        String(
          model.metadata
            .resultDescription || ''
        ).trim()
      );

    const articles =
      Array.isArray(
        model.metadata.resultArticles
      )
        ? model.metadata.resultArticles
        : [];

    if (articles.length === 0) {
      return description;
    }

    const listLines =
      articles.map(function (
        article,
        index
      ) {
        const articleName =
          QuestionJsonBuilder.formatText(
            article.name
          );

        const articleUrl =
          String(
            article.url || ''
          ).trim();

        return (
          `${index + 1}. ` +
          `[${articleName}]` +
          `(${articleUrl})`
        );
      });

    const list =
      '<list>\n' +
      listLines.join('\n') +
      '\n</list>';

    if (!description) {
      return list;
    }

    return [
      description,
      list
    ].join('\n');
  }

  static buildQuestions(model) {
    const questions =
      model.questions;

    const blockWeight =
      model.metadata.blockWeight;

    return questions.reduce(
      function (
        result,
        question
      ) {
        const questionId =
          question.id;

        const questionJson =
          QuestionJsonBuilder.build(
            question
          );

        if (questionId === 'q_1') {
          questionJson.blockWeight =
            blockWeight;
        }

        result[questionId] =
          questionJson;

        return result;
      },
      {}
    );
  }

}



/**
 * ==========================================
 * КОНЕЦ РАЗДЕЛА:
 * СБОРКА ИТОГОВОГО JSON
 * ==========================================
 */
