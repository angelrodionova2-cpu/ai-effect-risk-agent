const CONFIG = {
  SOURCE_SHEET: 'Ответы на форму (1)',
  RESULTS_SHEET: 'Результаты',
  RISKS_SHEET: 'Риски',
  MODEL: 'gemini-3.6-flash'
};


/**
 * ОСНОВНАЯ ФУНКЦИЯ.
 * Автоматически вызывается после отправки формы,
 * когда мы позже установим триггер.
 */
function onFormSubmit(e) {
  if (!e || !e.range) {
    throw new Error(
      'onFormSubmit запускается автоматически триггером. ' +
      'Для ручной проверки используйте testLatestResponse().'
    );
  }

  const sheet = e.range.getSheet();

  if (sheet.getName() !== CONFIG.SOURCE_SHEET) {
    return;
  }

  processSourceRow_(e.range.getRow(), true);
}


/**
 * ТЕСТ.
 * Обрабатывает последний уже существующий ответ формы.
 */
function testLatestResponse() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);

  if (!sheet) {
    throw new Error(
      `Не найден лист "${CONFIG.SOURCE_SHEET}"`
    );
  }

  const row = sheet.getLastRow();

  if (row < 2) {
    throw new Error('В форме пока нет ни одного ответа.');
  }

  processSourceRow_(row, false);
}


/**
 * ПЕРЕСБОРКА.
 * Удаляет старые расчёты и заново обрабатывает
 * все ответы формы.
 *
 * Используйте только при необходимости.
 */
function rebuildAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const source = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  const results = ss.getSheetByName(CONFIG.RESULTS_SHEET);
  const risks = ss.getSheetByName(CONFIG.RISKS_SHEET);

  if (!source || !results || !risks) {
    throw new Error('Не найден один из необходимых листов.');
  }

  if (results.getLastRow() > 1) {
    results
      .getRange(2, 1, results.getLastRow() - 1, results.getLastColumn())
      .clearContent();
  }

  if (risks.getLastRow() > 1) {
    risks
      .getRange(2, 1, risks.getLastRow() - 1, risks.getLastColumn())
      .clearContent();
  }

  for (let row = 2; row <= source.getLastRow(); row++) {
    processSourceRow_(row, false);
  }
}


/**
 * ОБРАБОТКА ОДНОГО ОТВЕТА.
 */
function processSourceRow_(rowNumber, shouldSendEmail = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const source = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  const resultsSheet = ss.getSheetByName(CONFIG.RESULTS_SHEET);
  const risksSheet = ss.getSheetByName(CONFIG.RISKS_SHEET);

  if (!source || !resultsSheet || !risksSheet) {
    throw new Error(
      'Проверьте наличие листов "Ответы на форму (1)", "Результаты" и "Риски".'
    );
  }

  const lastColumn = source.getLastColumn();

  const headers = source
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const values = source
    .getRange(rowNumber, 1, 1, lastColumn)
    .getValues()[0];


  // ------------------------------
  // 1. ИСХОДНЫЕ ДАННЫЕ
  // ------------------------------

  const timestamp = values[0];
  const email = getValueByHeader_(
  headers,
  values,
  'Адрес электронной почты'
);
  const processName =
    getValueByCode_(headers, values, 'PROCESS_NAME');

  const processGoal =
    getValueByCode_(headers, values, 'PROCESS_GOAL');

  const processInput =
    getValueByCode_(headers, values, 'PROCESS_INPUT');

  const processDescription =
    getValueByCode_(headers, values, 'PROCESS_DESCRIPTION');

  const processOutput =
    getValueByCode_(headers, values, 'PROCESS_OUTPUT');

  const readyCriteria =
    getValueByCode_(headers, values, 'READY_CRITERIA');


  // ------------------------------
  // 2. БЕЗ ИИ
  // ------------------------------

  const T0 = number_(
    getValueByCode_(headers, values, 'T0'),
    'T0'
  );

  const L0 = number_(
    getValueByCode_(headers, values, 'L0'),
    'L0'
  );

  const C0 = number_(
    getValueByCode_(headers, values, 'C0'),
    'C0'
  );

  const I0 = number_(
    getValueByCode_(headers, values, 'I0'),
    'I0'
  );

  const D0 =
    getValueByCode_(headers, values, 'D0');

  const q0Values =
    getValuesByCode_(headers, values, 'Q0')
      .map((v, i) => number_(v, `Q0-${i + 1}`));

  const Q0 = average_(q0Values);

  const W0 = number_(
    getValueByCode_(headers, values, 'W0'),
    'W0'
  );


  // ------------------------------
  // 3. С ИИ
  // ------------------------------

  const T1 = number_(
    getValueByCode_(headers, values, 'T1'),
    'T1'
  );

  const L1 = number_(
    getValueByCode_(headers, values, 'L1'),
    'L1'
  );

  const C1 = number_(
    getValueByCode_(headers, values, 'C1'),
    'C1'
  );

  const I1 = number_(
    getValueByCode_(headers, values, 'I1'),
    'I1'
  );

  const D1 =
    getValueByCode_(headers, values, 'D1');

  const q1Values =
    getValuesByCode_(headers, values, 'Q1')
      .map((v, i) => number_(v, `Q1-${i + 1}`));

  const Q1 = average_(q1Values);

  const W1 = number_(
    getValueByCode_(headers, values, 'W1'),
    'W1'
  );


  // ------------------------------
  // 4. РАСЧЁТ ЭФФЕКТИВНОСТИ
  // ------------------------------

  const deltaT = percentImprovement_(T0, T1);
  const deltaL = percentImprovement_(L0, L1);
  const deltaC = percentImprovement_(C0, C1);
  const deltaI = percentImprovement_(I0, I1);

  const deltaQ = Q1 - Q0;

  // Положительное значение = нагрузка уменьшилась.
  const deltaW = W0 - W1;


  // Баллы для IE.
  const ST = resourceScore_(T0, T1);
  const SL = resourceScore_(L0, L1);
  const SC = resourceScore_(C0, C1);
  const SI = resourceScore_(I0, I1);
  const SQ = qualityScore_(deltaQ);

  const IE = (ST + SL + SC + SI + SQ) / 20;

  const effectLevel = getEffectLevel_(IE);


  // ------------------------------
  // 5. КОНТЕКСТ ДЛЯ РИСКОВ
  // ------------------------------

  const aiUsage =
    getOptionalValueByCode_(headers, values, 'AI_USAGE');

  const dataTypes =
    getOptionalValueByCode_(headers, values, 'DATA_TYPES');

  const externalOutput =
    getOptionalValueByCode_(headers, values, 'EXTERNAL_OUTPUT');

  const humanReview =
    getOptionalValueByCode_(headers, values, 'HUMAN_REVIEW');

  const errorConsequence =
    getOptionalValueByCode_(headers, values, 'ERROR_CONSEQUENCE');

  const knownRisks =
    getOptionalValueByCode_(headers, values, 'KNOWN_RISKS');


  const context = {
    processName,
    processGoal,
    processInput,
    processDescription,
    processOutput,
    readyCriteria,

    aiUsage,
    dataTypes,
    externalOutput,
    humanReview,
    errorConsequence,
    knownRisks
  };


  // ------------------------------
  // 6. ИИ-АНАЛИЗ РИСКОВ
  // ------------------------------

  const risks = analyzeRisksWithGemini_(context);

  if (!risks.length) {
    throw new Error(
      'Gemini не вернул ни одного риска.'
    );
  }


  // ------------------------------
  // 7. РАСЧЁТ R
  // ------------------------------

  risks.forEach(risk => {
    risk.probability = clampInteger_(
      risk.probability,
      1,
      5
    );

    risk.impact = clampInteger_(
      risk.impact,
      1,
      5
    );

    // R считает КОД, а не нейросеть.
    risk.R =
      risk.probability *
      risk.impact;

    risk.riskLevel =
      getRiskLevel_(risk.R);

    risk.category =
      normalizeRiskCategory_(
        risk.category
      );
  });


  const Rmax = Math.max(
    ...risks.map(r => r.R)
  );

  const riskLevel =
    getRiskLevel_(Rmax);


  // ------------------------------
  // 8. МАТРИЦА ЭФФЕКТ–РИСК
  // ------------------------------

  const decision =
    getDecision_(
      effectLevel,
      riskLevel
    );


  // ------------------------------
  // 9. ЗАКЛЮЧЕНИЕ
  // ------------------------------

  const conclusion =
    buildConclusion_(
      IE,
      effectLevel,
      Rmax,
      riskLevel,
      decision,
      risks
    );


  // Если этот ответ уже тестировали,
  // удаляем старый результат, чтобы не было дублей.
  removeExistingRecords_(
    resultsSheet,
    risksSheet,
    timestamp,
    processName
  );


  // ------------------------------
  // 10. ЗАПИСЬ В РЕЗУЛЬТАТЫ
  // ------------------------------

  resultsSheet.appendRow([
    timestamp,                 // A Дата
    processName,               // B Процесс

    T0,                        // C
    T1,                        // D
    displayPercent_(deltaT),   // E

    L0,                        // F
    L1,                        // G
    displayPercent_(deltaL),   // H

    C0,                        // I
    C1,                        // J
    displayPercent_(deltaC),   // K

    I0,                        // L
    I1,                        // M
    displayPercent_(deltaI),   // N

    D0,                        // O
    D1,                        // P

    round_(Q0, 2),             // Q
    round_(Q1, 2),             // R
    round_(deltaQ, 2),         // S

    W0,                        // T
    W1,                        // U
    round_(deltaW, 2),         // V

    ST,                        // W
    SL,                        // X
    SC,                        // Y
    SI,                        // Z
    SQ,                        // AA

    round_(IE, 2),             // AB
    effectLevel,               // AC

    Rmax,                      // AD
    riskLevel,                 // AE
    decision,                  // AF
    conclusion                 // AG
  ]);


  // ------------------------------
  // 11. ЗАПИСЬ РИСКОВ
  // ------------------------------

  risks.forEach(risk => {
    risksSheet.appendRow([
      timestamp,
      processName,
      risk.risk_name || '',
      risk.category || '',
      risk.reason || '',
      risk.manifestation || '',
      risk.consequence || '',
      risk.probability,
      risk.impact,
      risk.R,
      risk.riskLevel,
      risk.mitigation || ''
    ]);
  });


  SpreadsheetApp.flush();
  if (shouldSendEmail && email) {
  sendResultEmail_({
    email: email,
    processName: processName,

    T0: T0,
    T1: T1,
    deltaT: deltaT,

    L0: L0,
    L1: L1,
    deltaL: deltaL,

    C0: C0,
    C1: C1,
    deltaC: deltaC,

    I0: I0,
    I1: I1,
    deltaI: deltaI,

    D0: D0,
    D1: D1,

    Q0: Q0,
    Q1: Q1,
    deltaQ: deltaQ,

    W0: W0,
    W1: W1,
    deltaW: deltaW,

    IE: IE,
    effectLevel: effectLevel,

    Rmax: Rmax,
    riskLevel: riskLevel,

    decision: decision,
    conclusion: conclusion,

    risks: risks
  });
}
  Logger.log(
    `Готово. Процесс: ${processName}. ` +
    `IE=${round_(IE, 2)}, ` +
    `Rmax=${Rmax}, ` +
    `${decision}`
  );
}


/**
 * ЗАПРОС К GEMINI.
 */
function analyzeRisksWithGemini_(context) {
  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error(
      'Не найден GEMINI_API_KEY. ' +
      'Добавьте его в свойства скрипта.'
    );
  }

  const prompt = `
Ты — аналитик рисков применения генеративного искусственного интеллекта в бизнес-процессах.

Не оценивай "ИИ вообще".
Оценивай только конкретный бизнес-процесс и конкретный сценарий применения ИИ, описанные ниже.

ЗАДАЧА:
Выяви от 3 до 6 наиболее существенных рисков применения ИИ.

Используй только четыре категории:

LEG — юридические и репутационные риски;
ORG — организационные и кадровые риски;
QLT — качество результата и ошибки модели;
BIZ — иные существенные бизнес-риски.

Для каждого риска определи:

- название риска;
- категорию;
- причину;
- возможное проявление;
- возможное последствие;
- вероятность P от 1 до 5;
- влияние V от 1 до 5;
- конкретную меру управления.

ШКАЛА ВЕРОЯТНОСТИ P:

1 — маловероятно;
2 — скорее редко;
3 — возможно;
4 — вероятно;
5 — очень вероятно.

ШКАЛА ВЛИЯНИЯ V:

1 — несущественное влияние;
2 — слабое влияние;
3 — заметное влияние;
4 — серьёзное влияние;
5 — критическое влияние.

ВАЖНО:

1. Не рассчитывай R = P × V. Это сделает программа.
2. Не завышай риски автоматически.
3. Не придумывай факты, которых нет во входных данных.
4. Если пользователь указал существующую человеческую проверку, учитывай её как фактор снижения риска.
5. Если результат может попасть клиенту или во внешнюю коммуникацию, учитывай это при оценке влияния.
6. Если передаются внутренние, персональные или коммерчески чувствительные данные, оцени соответствующий риск.
7. Если пользователь уже наблюдал конкретную проблему, учитывай это при оценке вероятности.
8. Формулируй меры управления практически и кратко.
9. Строго отделяй факты из анкеты от предположений.
10. Не утверждай наличие персональных данных, коммерческой тайны,
NDA, договорных ограничений, штрафов, законодательных нарушений,
сертификаций или иных обстоятельств, если они прямо не указаны
во входных данных.
11. Если возможное последствие является только предположением,
формулируй его как возможность, а не как установленный факт.
12. Для оценки риска используй только информацию, содержащуюся
во входных данных, и общие свойства описанного способа применения ИИ.
13. Не добавляй специфические характеристики компании,
клиента, договора, продукта или ИИ-сервиса, которых нет в анкете.
14. Если пользователь указал только общедоступную информацию
и не выбрал внутренние, персональные или коммерчески чувствительные
данные, не создавай отдельный риск утечки конфиденциальной информации
только на основании того, что в ИИ загружаются файлы или изображения.
15. Не заменяй явно выбранные пользователем категории данных
предположениями о том, что файлы "могут случайно содержать"
другие типы информации. Оценивай фактически указанный сценарий.

ДАННЫЕ ПРОЦЕССА:

Название:
${context.processName}

Цель:
${context.processGoal}

Описание процесса:
${context.processDescription}

Исходные данные:
${context.processInput}

Результат:
${context.processOutput}

Критерий готовности:
${context.readyCriteria}

Как используется ИИ:
${context.aiUsage}

Какие данные передаются ИИ:
${context.dataTypes}

Может ли результат использоваться во внешней коммуникации:
${context.externalOutput}

Как организована проверка человеком:
${context.humanReview}

Возможные последствия серьёзной ошибки:
${context.errorConsequence}

Уже замеченные проблемы и опасения:
${context.knownRisks || 'Не указаны'}
`;


  const schema = {
    type: 'object',
    properties: {
      risks: {
        type: 'array',
        minItems: 3,
        maxItems: 6,
        items: {
          type: 'object',
          properties: {
            risk_name: {
              type: 'string'
            },
            category: {
              type: 'string',
              enum: [
                'LEG',
                'ORG',
                'QLT',
                'BIZ'
              ]
            },
            reason: {
              type: 'string'
            },
            manifestation: {
              type: 'string'
            },
            consequence: {
              type: 'string'
            },
            probability: {
              type: 'integer',
              minimum: 1,
              maximum: 5
            },
            impact: {
              type: 'integer',
              minimum: 1,
              maximum: 5
            },
            mitigation: {
              type: 'string'
            }
          },

          required: [
            'risk_name',
            'category',
            'reason',
            'manifestation',
            'consequence',
            'probability',
            'impact',
            'mitigation'
          ]
        }
      }
    },

    required: ['risks']
  };


  const payload = {
    contents: [
      {
        role: 'user',
        parts: [
          {
            text: prompt
          }
        ]
      }
    ],

    generationConfig: {
      responseMimeType: 'application/json',
      responseJsonSchema: schema,
      maxOutputTokens: 3000
    }
  };


  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    CONFIG.MODEL +
    ':generateContent?key=' +
    encodeURIComponent(apiKey);


  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: 'post',

        contentType:
          'application/json',

        payload:
          JSON.stringify(payload),

        muteHttpExceptions:
          true
      }
    );


  const status =
    response.getResponseCode();

  const body =
    response.getContentText();


  if (status < 200 || status >= 300) {
    throw new Error(
      `Ошибка Gemini API ${status}: ${body}`
    );
  }


  const data =
    JSON.parse(body);


  if (
    !data.candidates ||
    !data.candidates.length ||
    !data.candidates[0].content ||
    !data.candidates[0].content.parts
  ) {
    throw new Error(
      'Gemini вернул неожиданный ответ: ' +
      body
    );
  }


  const text =
    data.candidates[0]
      .content
      .parts
      .map(part => part.text || '')
      .join('');


  const parsed =
    JSON.parse(text);


  if (
    !parsed.risks ||
    !Array.isArray(parsed.risks)
  ) {
    throw new Error(
      'Gemini не вернул массив risks.'
    );
  }


  return parsed.risks;
}


/**
 * Поиск одного поля формы по техническому коду.
 */
function getValueByCode_(headers, values, code) {
  const tag = `[${code}]`;

  const index =
    headers.findIndex(
      h => String(h).includes(tag)
    );

  if (index === -1) {
    throw new Error(
      `Не найден вопрос с кодом ${tag}`
    );
  }

  return values[index];
}


/**
 * Поиск необязательного поля.
 */
function getOptionalValueByCode_(headers, values, code) {
  const tag = `[${code}]`;

  const index =
    headers.findIndex(
      h => String(h).includes(tag)
    );

  if (index === -1) {
    return '';
  }

  return values[index] || '';
}


/**
 * Поиск нескольких столбцов с одним кодом.
 * Используется для сеток Q0 и Q1.
 */
function getValuesByCode_(headers, values, code) {
  const tag = `[${code}]`;

  const result = [];

  headers.forEach((header, index) => {
    if (
      String(header).includes(tag)
    ) {
      result.push(values[index]);
    }
  });

  if (!result.length) {
    throw new Error(
      `Не найдены столбцы ${tag}`
    );
  }

  return result;
}


/**
 * Преобразование в число.
 */
function number_(value, name) {
  if (
    typeof value === 'number' &&
    !isNaN(value)
  ) {
    return value;
  }

  const prepared =
    String(value)
      .trim()
      .replace(/\s/g, '')
      .replace(',', '.');

  const number =
    Number(prepared);

  if (isNaN(number)) {
    throw new Error(
      `${name}: ожидается число, получено "${value}"`
    );
  }

  return number;
}


/**
 * Среднее арифметическое.
 */
function average_(array) {
  if (!array.length) {
    return 0;
  }

  return (
    array.reduce(
      (sum, value) => sum + value,
      0
    ) / array.length
  );
}


/**
 * Процент улучшения.
 *
 * Положительное значение = улучшение.
 */
function percentImprovement_(before, after) {
  if (before === 0) {
    if (after === 0) {
      return 0;
    }

    return null;
  }

  return (
    (before - after) /
    before *
    100
  );
}


/**
 * Балл T/L/C/I для IE.
 *
 * >20%     = 4
 * 10–20%   = 3
 * 0–10%    = 2
 * до -10%  = 1
 * ниже -10 = 0
 *
 * При нулевой базе:
 * 0 -> 0 считается отсутствием изменения;
 * 0 -> >0 считается ухудшением.
 */
function resourceScore_(before, after) {
  if (before === 0) {
    return after === 0 ? 2 : 0;
  }

  const change =
    percentImprovement_(
      before,
      after
    );

  if (change > 20) {
    return 4;
  }

  if (change >= 10) {
    return 3;
  }

  if (change >= 0) {
    return 2;
  }

  if (change >= -10) {
    return 1;
  }

  return 0;
}


/**
 * Балл качества.
 */
function qualityScore_(change) {
  if (change > 0.5) {
    return 4;
  }

  if (change > 0) {
    return 3;
  }

  if (change === 0) {
    return 2;
  }

  if (change >= -0.5) {
    return 1;
  }

  return 0;
}


/**
 * Уровень эффекта по IE.
 */
function getEffectLevel_(IE) {
  if (IE >= 0.80) {
    return 'Высокий';
  }

  if (IE >= 0.60) {
    return 'Заметный';
  }

  if (IE >= 0.40) {
    return 'Умеренный';
  }

  return 'Слабый или противоречивый';
}


/**
 * Уровень риска по R.
 */
function getRiskLevel_(R) {
  if (R <= 4) {
    return 'Низкий';
  }

  if (R <= 9) {
    return 'Умеренный';
  }

  if (R <= 16) {
    return 'Значимый';
  }

  return 'Высокий';
}


/**
 * Матрица "эффект–риск".
 */
function getDecision_(effectLevel, riskLevel) {
  const goodEffect =
    effectLevel === 'Высокий' ||
    effectLevel === 'Заметный';

  const manageableRisk =
    riskLevel === 'Низкий' ||
    riskLevel === 'Умеренный';


  if (
    goodEffect &&
    manageableRisk
  ) {
    return (
      'Тип 1 — масштабировать как устойчивый сценарий ' +
      'с базовым контролем'
    );
  }


  if (
    goodEffect &&
    !manageableRisk
  ) {
    return (
      'Тип 2 — масштабировать только с регламентом ' +
      'и мерами контроля'
    );
  }


  if (
    !goodEffect &&
    manageableRisk
  ) {
    return (
      'Тип 3 — доработать сценарий применения'
    );
  }


  return (
    'Тип 4 — не масштабировать в текущем виде'
  );
}


/**
 * Автоматическое управленческое заключение.
 */
function buildConclusion_(
  IE,
  effectLevel,
  Rmax,
  riskLevel,
  decision,
  risks
) {
  const highestRisks =
    risks
      .filter(r => r.R === Rmax)
      .slice(0, 2);


  const riskNames =
    highestRisks
      .map(r => r.risk_name)
      .join('; ');


  const measures = [
    ...new Set(
      highestRisks
        .map(r => r.mitigation)
        .filter(Boolean)
    )
  ].join('; ');


  return (
    `Применение ИИ показывает ${effectLevel.toLowerCase()} ` +
    `эффект (IE = ${round_(IE, 2)}). ` +
    `Максимальный риск Rmax = ${Rmax}, ` +
    `уровень риска — ${riskLevel.toLowerCase()}. ` +
    `Наиболее существенный риск: ${riskNames}. ` +
    `${decision}. ` +
    `Ключевые меры управления: ${measures}.`
  );
}


/**
 * Удаляет предыдущий расчёт для того же ответа.
 * Позволяет повторно запускать тест без дублей.
 */
function removeExistingRecords_(
  resultsSheet,
  risksSheet,
  timestamp,
  processName
) {
  removeRowsByKey_(
    resultsSheet,
    timestamp,
    processName
  );

  removeRowsByKey_(
    risksSheet,
    timestamp,
    processName
  );
}


function removeRowsByKey_(
  sheet,
  timestamp,
  processName
) {
  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        2
      )
      .getValues();


  for (
    let i = data.length - 1;
    i >= 0;
    i--
  ) {
    const rowTimestamp =
      data[i][0];

    const rowProcess =
      data[i][1];


    if (
      sameTimestamp_(
        rowTimestamp,
        timestamp
      ) &&
      String(rowProcess) ===
      String(processName)
    ) {
      sheet.deleteRow(i + 2);
    }
  }
}


/**
 * Сравнение дат.
 */
function sameTimestamp_(a, b) {
  if (
    a instanceof Date &&
    b instanceof Date
  ) {
    return (
      a.getTime() ===
      b.getTime()
    );
  }

  return (
    String(a) ===
    String(b)
  );
}


/**
 * Категория риска.
 */
function normalizeRiskCategory_(value) {
  const allowed =
    ['LEG', 'ORG', 'QLT', 'BIZ'];

  const normalized =
    String(value)
      .trim()
      .toUpperCase();

  return allowed.includes(normalized)
    ? normalized
    : 'BIZ';
}


/**
 * Целое число в допустимом диапазоне.
 */
function clampInteger_(
  value,
  min,
  max
) {
  const n =
    Math.round(
      Number(value)
    );

  if (isNaN(n)) {
    return min;
  }

  return Math.max(
    min,
    Math.min(max, n)
  );
}


/**
 * Округление.
 */
function round_(value, digits) {
  const factor =
    Math.pow(10, digits);

  return (
    Math.round(
      (value + Number.EPSILON) *
      factor
    ) / factor
  );
}


/**
 * Для вывода процентного изменения.
 */
function displayPercent_(value) {
  if (value === null) {
    return 'н/р';
  }

  return round_(value, 2);
}
/**
 * Получение значения по точному заголовку столбца.
 */
function getValueByHeader_(headers, values, headerName) {
  const index = headers.findIndex(
    h => String(h).trim() === headerName
  );

  if (index === -1) {
    return '';
  }

  return String(values[index] || '').trim();
}


/**
 * Отправляет пользователю готовый отчёт.
 */
function sendResultEmail_(report) {

  if (!report.email) {
    return;
  }

  const remainingQuota =
    MailApp.getRemainingDailyQuota();

  if (remainingQuota < 1) {
    throw new Error(
      'Исчерпан дневной лимит отправки email.'
    );
  }


  const categoryNames = {
    LEG: 'Юридические и репутационные',
    ORG: 'Организационные и кадровые',
    QLT: 'Качество результата и ошибки модели',
    BIZ: 'Иные бизнес-риски'
  };


  const risksHtml = report.risks
    .map((risk, index) => {

      const category =
        categoryNames[risk.category] ||
        risk.category;

      return `
        <div style="
          margin: 0 0 16px 0;
          padding: 16px;
          border: 1px solid #e5e7eb;
          border-radius: 10px;
        ">

          <div style="
            font-size: 16px;
            font-weight: 700;
            margin-bottom: 6px;
          ">
            ${index + 1}. ${escapeHtml_(risk.risk_name)}
          </div>

          <div style="
            font-size: 13px;
            color: #6b7280;
            margin-bottom: 10px;
          ">
            ${escapeHtml_(category)}
            · P = ${risk.probability}
            · V = ${risk.impact}
            · R = ${risk.R}
            · ${escapeHtml_(risk.riskLevel)}
          </div>

          <div style="margin-bottom: 6px;">
            <strong>Причина:</strong>
            ${escapeHtml_(risk.reason)}
          </div>

          <div style="margin-bottom: 6px;">
            <strong>Возможное проявление:</strong>
            ${escapeHtml_(risk.manifestation)}
          </div>

          <div style="margin-bottom: 6px;">
            <strong>Возможное последствие:</strong>
            ${escapeHtml_(risk.consequence)}
          </div>

          <div>
            <strong>Мера управления:</strong>
            ${escapeHtml_(risk.mitigation)}
          </div>

        </div>
      `;
    })
    .join('');


  const htmlBody = `
    <div style="
      max-width: 760px;
      margin: 0 auto;
      padding: 24px;
      font-family: Arial, Helvetica, sans-serif;
      color: #111827;
      line-height: 1.5;
    ">

      <h1 style="
        font-size: 24px;
        margin: 0 0 8px 0;
      ">
        Результат оценки применения ИИ
      </h1>

      <p style="
        color: #6b7280;
        margin-top: 0;
      ">
        Оценка эффективности и рисков конкретного бизнес-процесса
      </p>


      <div style="
        padding: 16px;
        background: #f9fafb;
        border-radius: 10px;
        margin: 24px 0;
      ">

        <div style="
          font-size: 13px;
          color: #6b7280;
        ">
          Анализируемый процесс
        </div>

        <div style="
          font-size: 20px;
          font-weight: 700;
        ">
          ${escapeHtml_(report.processName)}
        </div>

      </div>


      <h2 style="font-size: 19px;">
        1. Эффективность
      </h2>


      <table
        cellpadding="8"
        cellspacing="0"
        style="
          width: 100%;
          border-collapse: collapse;
          margin-bottom: 20px;
        "
      >

        <tr>
          <th align="left">Показатель</th>
          <th align="right">Без ИИ</th>
          <th align="right">С ИИ</th>
          <th align="right">Изменение</th>
        </tr>

        ${emailMetricRow_(
          'Активное время',
          emailNumber_(report.T0) + ' ч',
          emailNumber_(report.T1) + ' ч',
          emailPercent_(report.deltaT)
        )}

        ${emailMetricRow_(
          'Трудоёмкость',
          emailNumber_(report.L0) + ' чел.-ч',
          emailNumber_(report.L1) + ' чел.-ч',
          emailPercent_(report.deltaL)
        )}

        ${emailMetricRow_(
          'Оценочная стоимость',
          emailNumber_(report.C0) + ' ₽',
          emailNumber_(report.C1) + ' ₽',
          emailPercent_(report.deltaC)
        )}

        ${emailMetricRow_(
          'Существенные итерации',
          emailNumber_(report.I0, 0),
          emailNumber_(report.I1, 0),
          emailPercent_(report.deltaI)
        )}

        ${emailMetricRow_(
          'Качество',
          emailNumber_(report.Q0) + ' / 5',
          emailNumber_(report.Q1) + ' / 5',
          signedNumber_(report.deltaQ) + ' балла'
        )}

        ${emailMetricRow_(
          'Субъективная нагрузка',
          emailNumber_(report.W0) + ' / 5',
          emailNumber_(report.W1) + ' / 5',
          'снижение на ' +
            emailNumber_(report.deltaW) +
            ' балла'
        )}

        ${emailMetricRow_(
          'Выполнение в срок',
          escapeHtml_(report.D0),
          escapeHtml_(report.D1),
          ''
        )}

      </table>


      <div style="
        padding: 16px;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        margin-bottom: 24px;
      ">

        <strong>Индекс эффективности IE:</strong>
        ${emailNumber_(report.IE)}

        <br>

        <strong>Уровень эффекта:</strong>
        ${escapeHtml_(report.effectLevel)}

      </div>


      <h2 style="font-size: 19px;">
        2. Риски
      </h2>

      ${risksHtml}


      <div style="
        padding: 16px;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        margin: 20px 0 24px 0;
      ">

        <strong>Максимальный риск Rmax:</strong>
        ${report.Rmax}

        <br>

        <strong>Уровень риска:</strong>
        ${escapeHtml_(report.riskLevel)}

      </div>


      <h2 style="font-size: 19px;">
        3. Управленческое решение
      </h2>

      <div style="
        padding: 18px;
        background: #f9fafb;
        border-radius: 10px;
        margin-bottom: 16px;
      ">

        <div style="
          font-size: 18px;
          font-weight: 700;
        ">
          ${escapeHtml_(report.decision)}
        </div>

      </div>


      <p>
        ${escapeHtml_(report.conclusion)}
      </p>


      <hr style="
        border: 0;
        border-top: 1px solid #e5e7eb;
        margin: 28px 0 18px 0;
      ">


      <p style="
        font-size: 12px;
        color: #6b7280;
      ">
        Эффективность рассчитывается автоматически на основании
        введённых показателей. Риски выявляются с использованием
        ИИ на основании информации, указанной в анкете.
        Балл риска рассчитывается программно по формуле R = P × V.
        Итог является аналитической рекомендацией и предполагает
        проверку ответственным специалистом.
      </p>

    </div>
  `;


  const textBody =
    'Результат оценки применения ИИ\n\n' +

    'Процесс: ' +
    report.processName +
    '\n\n' +

    'Уровень эффекта: ' +
    report.effectLevel +
    '\n' +

    'IE: ' +
    emailNumber_(report.IE) +
    '\n\n' +

    'Максимальный риск Rmax: ' +
    report.Rmax +
    '\n' +

    'Уровень риска: ' +
    report.riskLevel +
    '\n\n' +

    'Решение:\n' +
    report.decision +
    '\n\n' +

    'Заключение:\n' +
    report.conclusion;


  MailApp.sendEmail({
    to: report.email,

    subject:
      'Результат оценки применения ИИ — ' +
      report.processName,

    body: textBody,

    htmlBody: htmlBody,

    name:
      'AI Effect Risk Agent'
  });
}


/**
 * Строка таблицы показателей.
 */
function emailMetricRow_(
  name,
  before,
  after,
  change
) {

  return `
    <tr style="
      border-top: 1px solid #e5e7eb;
    ">

      <td>
        ${escapeHtml_(name)}
      </td>

      <td align="right">
        ${before}
      </td>

      <td align="right">
        ${after}
      </td>

      <td align="right">
        ${change}
      </td>

    </tr>
  `;
}


/**
 * Формат числа для письма.
 */
function emailNumber_(value, digits = 2) {

  if (
    value === null ||
    value === undefined ||
    value === '' ||
    isNaN(Number(value))
  ) {
    return '—';
  }

  return String(
    round_(Number(value), digits)
  ).replace('.', ',');
}


/**
 * Формат процента.
 */
function emailPercent_(value) {

  if (value === null) {
    return 'н/р';
  }

  return (
    emailNumber_(value) +
    '%'
  );
}


/**
 * Число со знаком + / -.
 */
function signedNumber_(value) {

  const number =
    Number(value);

  if (isNaN(number)) {
    return '—';
  }

  const formatted =
    emailNumber_(number);

  if (number > 0) {
    return '+' + formatted;
  }

  return formatted;
}


/**
 * Защита HTML.
 */
function escapeHtml_(value) {

  return String(
    value === null ||
    value === undefined
      ? ''
      : value
  )
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}


/**
 * Ручной тест отправки письма.
 * Обрабатывает последний ответ и ОТПРАВЛЯЕТ email.
 */
function testLatestResponseWithEmail() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      CONFIG.SOURCE_SHEET
    );

  if (!sheet) {
    throw new Error(
      `Не найден лист "${CONFIG.SOURCE_SHEET}"`
    );
  }

  const row =
    sheet.getLastRow();

  if (row < 2) {
    throw new Error(
      'Нет ответов формы.'
    );
  }

  processSourceRow_(
    row,
    true
  );
}
