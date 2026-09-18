/**
 * Calamity City — бэкенд билдера отрядов на Google Apps Script.
 *
 * Что делает этот файл:
 *  - при первом запуске сам создаёт на этой таблице вкладки с игровыми данными
 *    (оружие, снаряжение, способности, архетипы, роли, спецправила, глоссарий)
 *    и заполняет их стартовыми значениями;
 *  - отдаёт эти данные сайту-билдеру в виде JSON (doGet);
 *  - хранит сохранённые отряды игроков во вкладке «Отряды» и отдаёт/сохраняет/
 *    удаляет их по имени + коду доступа (doPost).
 *
 * УСТАНОВКА — см. README.md в корне проекта. Коротко:
 *  1. Создайте новую Google Таблицу.
 *  2. Расширения → Apps Script, вставьте сюда весь этот файл целиком (замените
 *     содержимое Code.gs), сохраните.
 *  3. Разверните → Новое развёртывание → тип «Веб-приложение»:
 *       Выполнять как: я
 *       У кого есть доступ: Все (Anyone)
 *     Скопируйте появившийся URL — это и есть API_URL для index.html.
 *  4. Обновите таблицу в браузере — сверху появится меню «Calamity City»,
 *     нажмите «Создать/обновить вкладки с начальными данными» один раз.
 */

/* ============================================================ КОНФИГ ==== */

var SHEET_WEAPONS = 'Оружие';
var SHEET_EQUIPMENT = 'Снаряжение';
var SHEET_ABILITIES = 'Способности';
var SHEET_ARCHETYPES = 'Архетипы';
var SHEET_ROLES = 'Роли';
var SHEET_SPECIAL_RULES = 'Спецправила';
var SHEET_GLOSSARY = 'Глоссарий';
var SHEET_SQUADS = 'Отряды';

/* ============================================================== МЕНЮ ==== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calamity City')
    .addItem('Создать/обновить вкладки с начальными данными', 'seedAll')
    .addToUi();
}

/* ===================================================== СТАРТОВЫЕ ДАННЫЕ ==
   Ровно то же самое, что раньше лежало прямо в коде сайта (DEFAULT_DATA).
   Меняются они теперь через сами вкладки таблицы, а не здесь — но если
   вкладка пустая, seedAll() один раз заполнит её отсюда. */

var SEED_DATA = {
  archetypes: {
    organization: {
      id: 'organization', name: 'Организация',
      blurb: 'Профессиональные оперативники: военные, агенты, наёмники, специалисты — полагаются на подготовку и снаряжение, а не на сверхспособности.',
      archetypeCategories: ['Общие', 'Стратегия', 'Беспощадность', 'Точность'],
      specialRules: [
        { name: 'Сосредоточенные на миссии', text: 'Когда персонаж Организации впервые за раунд успешно проходит проверку, связанную с миссией — отряд получает 1 Очко Мотивации.' },
        { name: 'Специалисты широкого профиля', text: 'Командир и Ветераны Организации могут выбрать ещё одну категорию способностей; для них она считается способностью архетипа (билдер это не считает автоматически — выбирайте вручную).' },
        { name: 'Цепочка командования', text: 'Пока Ветеран Организации в пределах 8" от Командира — он снижает сложность простых проверок Интеллекта и Воли на 1. Пока Новобранец в пределах 8" от Ветерана — то же самое.' }
      ],
      roles: {
        commander: { id: 'commander', name: 'Командир', tier: 'senior', maxCount: 1, sizeMin: 1, sizeMax: 2, defaultSize: 2, cost: 10, level: 8, statIncreases: 5, archetypeAbilitySlots: 2, anyAbilitySlots: 1 },
        veteran: { id: 'veteran', name: 'Ветеран', tier: 'senior', maxCount: 3, sizeMin: 1, sizeMax: 2, defaultSize: 2, cost: 7, level: 6, statIncreases: 4, archetypeAbilitySlots: 2, anyAbilitySlots: 0 },
        heavyVeteran: { id: 'heavyVeteran', name: 'Тяжёлый Ветеран', tier: 'senior', maxCount: 2, sizeMin: 2, sizeMax: 3, defaultSize: 3, cost: 8, level: 6, statIncreases: 3, archetypeAbilitySlots: 2, anyAbilitySlots: 0, builtinArmor: 1, builtinHealth: 2, mightBaseline: 'd8', speedCapped: true },
        recruit: { id: 'recruit', name: 'Новобранец', tier: 'novice', maxCount: 6, sizeMin: 1, sizeMax: 2, defaultSize: 2, cost: 3, level: 3, statIncreases: 2, archetypeAbilitySlots: 0, anyAbilitySlots: 0 }
      }
    },
    gifted: {
      id: 'gifted', name: 'Одарённые',
      blurb: 'Мутанты, маги, вампиры, культисты, сверхлюди — их силы выходят далеко за пределы нормы человека.',
      archetypeCategories: ['Общие', 'Стратегия', 'Вдохновение', 'Больше чем человек', 'Духовные силы', 'Манифестация'],
      specialRules: [
        { name: 'Особая сила', text: 'В начале фазы действия Командир/Ветеран может использовать способность Духовной силы или Манифестации без траты Очков Мотивации (без урона врагам; при переменной стоимости — как минимальная).' },
        { name: 'Сверхъестественная связь', text: 'Когда Командир/Ветеран впервые за раунд тратит Очки Мотивации — дружественный Новобранец в 12" получает Быстрое Ожидание.' },
        { name: 'Пробуждение', text: 'Когда Командир/Ветеран в свою активацию успешно применяет Сверхъестественную Силу — получает быстрое действие Взаимодействия.' },
        { name: '«Больше, чем человек» (порог)', text: 'При 3+ способностях этой категории у персонажа — не может Взаимодействовать с нейтралами через проверки Воли. При 4+ — также не может Взаимодействовать с объектами через проверки Интеллекта.' }
      ],
      roles: {
        commander: { id: 'commander', name: 'Командир', tier: 'senior', maxCount: 1, sizeMin: 1, sizeMax: 2, defaultSize: 2, cost: 15, level: 9, statIncreases: 5, archetypeAbilitySlots: 3, anyAbilitySlots: 1, archetypeCategories: ['Общие', 'Стратегия', 'Вдохновение', 'Больше чем человек', 'Духовные силы', 'Манифестация'] },
        veteran: { id: 'veteran', name: 'Ветеран', tier: 'senior', maxCount: 3, sizeMin: 1, sizeMax: 2, defaultSize: 2, cost: 12, level: 7, statIncreases: 4, archetypeAbilitySlots: 2, anyAbilitySlots: 1, archetypeCategories: ['Общие', 'Стратегия', 'Вдохновение', 'Больше чем человек', 'Духовные силы', 'Манифестация'] },
        heavyRecruit: { id: 'heavyRecruit', name: 'Тяжёлый Новобранец', tier: 'novice', maxCount: 2, sizeMin: 3, sizeMax: 3, defaultSize: 3, cost: 7, level: 5, statIncreases: 3, archetypeAbilitySlots: 1, anyAbilitySlots: 0, archetypeCategories: ['Общие', 'Проворство', 'Мощь', 'Больше чем человек'], note: 'Дополнительно получает 1 бесплатную встроенную способность категории «Больше чем человек» — билдер её не считает в лимит слотов, впишите вручную на карточке.' },
        recruit: { id: 'recruit', name: 'Новобранец', tier: 'novice', maxCount: 3, sizeMin: 1, sizeMax: 2, defaultSize: 2, cost: 5, level: 4, statIncreases: 2, archetypeAbilitySlots: 1, anyAbilitySlots: 0, archetypeCategories: ['Общие', 'Проворство', 'Мощь', 'Больше чем человек'] }
      }
    },
    heavy: {
      id: 'heavy', name: 'Тяжёлый отряд',
      blurb: 'Тяжелобронированные солдаты, роботы, мехи и могущественные потусторонние существа — внушительный размер и разрушительная сила.',
      archetypeCategories: ['Общие', 'Мощь', 'Беспощадность', 'Точность', 'Больше чем человек'],
      sizeBonusRule: { size: 3, armor: 1, health: 1 },
      specialRules: [
        { name: 'Шок и трепет', text: 'В начале фазы действия отряд получает доп. Очко Мотивации. Тратить Очки Мотивации могут только персонажи размера 3.' },
        { name: 'За гранью возможного', text: 'Раз за раунд, в начале активации любого персонажа размера 3 — он может получить быстрое действие.' },
        { name: 'Неостановимая сила', text: 'Персонажи размера 3 получают встроенную Броню 1 и +1 Здоровье (билдер считает это автоматически по размеру).' },
        { name: '«Больше, чем человек» (порог)', text: 'При 3+ способностях этой категории у персонажа — не может Взаимодействовать с нейтралами через проверки Воли. При 4+ — также не может Взаимодействовать с объектами через проверки Интеллекта.' }
      ],
      roles: {
        commander: { id: 'commander', name: 'Командир', tier: 'senior', maxCount: 1, sizeMin: 3, sizeMax: 3, defaultSize: 3, cost: 16, level: 9, statIncreases: 5, archetypeAbilitySlots: 3, anyAbilitySlots: 1 },
        veteran: { id: 'veteran', name: 'Ветеран', tier: 'senior', maxCount: 3, sizeMin: 3, sizeMax: 3, defaultSize: 3, cost: 13, level: 7, statIncreases: 4, archetypeAbilitySlots: 2, anyAbilitySlots: 1 },
        heavyRecruit: { id: 'heavyRecruit', name: 'Тяжёлый Новобранец', tier: 'novice', maxCount: 3, sizeMin: 3, sizeMax: 3, defaultSize: 3, cost: 8, level: 5, statIncreases: 3, archetypeAbilitySlots: 1, anyAbilitySlots: 0 },
        recruit: { id: 'recruit', name: 'Новобранец', tier: 'novice', maxCount: 3, sizeMin: 1, sizeMax: 2, defaultSize: 2, cost: 5, level: 4, statIncreases: 2, archetypeAbilitySlots: 1, anyAbilitySlots: 0 }
      }
    }
  },

  abilities: {
    ambidexter: { name: 'Амбидекстер', category: 'Общие', text: 'Если персонаж одновременно вооружён двумя оружиями ближнего боя или двумя стрелковыми без пометки Двуручное и одно из них Лёгкое — при стандартном действии Сражения/Стрельбы: получает быстрое действие того же типа, либо урон используемого оружия +1 (если дальность второго не меньше).' },
    zdorovyak: { name: 'Здоровяк', category: 'Общие', text: 'Персонаж получает +1 к Здоровью и +1 слот инвентаря.' },
    stubborn: { name: 'Упорный', category: 'Общие', text: 'Когда персонаж тратит Очко Мотивации для понижения сложности проверки и проверка всё равно проваливается — Очко Мотивации восстанавливается.' },
    resilient: { name: 'Живучий', category: 'Общие', text: 'В конце активации при Здоровье ≤3 (и неполном) — проверка Воли(4): успех восстанавливает 1 Здоровье.' },
    lucky: { name: 'Везучий', category: 'Общие', text: 'Когда для успеха не хватает ровно 1 — можно потратить Очко Мотивации, снизив сложность на 1.' },
    hyperactive: { name: 'Гиперактивный', category: 'Общие', text: 'До двух разных эффектов Мотивации за активацию вместо одного (не одинаковых).' },

    acrobatics: { name: 'Акробатические движения', category: 'Проворство', text: 'Не получает урона от падения при добровольном перемещении. Особое Взаимодействие дальностью 2 на элемент террейна своего размера или менее: проверка Ловкости(4), успех — переместиться в контакт с элементом, супер-успех — в пределах 3".' },
    sprint: { name: 'Спринт', category: 'Проворство', text: 'После стандартного Перемещения — быстрое действие Перемещения.' },
    charge_attack: { name: 'Атака на ходу', category: 'Проворство', text: 'После стандартного Рывка — быстрое действие Сражения.' },
    run_and_gun: { name: 'Беги и стреляй', category: 'Проворство', text: 'После стандартного действия Стрельбы — быстрое действие Рывка.' },
    lightning_reaction: { name: 'Молниеносная реакция', category: 'Проворство', text: 'При действии Ожидания (стандартном или быстром) можно получить статус Ожидание(Рывок) / Быстрое Ожидание(Рывок).' },

    iron_fist: { name: 'Железный кулак', category: 'Мощь', text: 'Безоружная атака становится: ближний бой, дальность 1, урон 2, Бронебойность(1), Пролом(4+), Лёгкое. Персонаж всегда вооружён двумя такими «оружиями».' },
    born_wrestler: { name: 'Рестлер по призванию', category: 'Мощь', text: '+1 слот инвентаря. Особое Взаимодействие дальностью 1 на врага равного/меньшего размера: встречная проверка Силы (+2 если цель меньше) — успех: Бросок цели на дистанцию = максимум Силы; супер-успех: урон столкновения +1.' },
    ramming_blow: { name: 'Таранный удар', category: 'Мощь', text: 'Если атака Сражения сняла ≥2 Здоровья — можно совершить Бросок цели на 3".' },
    crushing_blows: { name: 'Дробящие удары', category: 'Мощь', text: 'Атака Сражения оружием с Требует Силу, снявшая ≥2 Здоровья — цель проходит проверку Воли(4) или получает Оглушение.' },

    tactical_advantage: { name: 'Тактическое преимущество', category: 'Стратегия', text: 'Взаимодействие дальностью 6" на дружественного: проверка Интеллекта(4) — успех: Ожидание; супер-успех: ещё один дружественный в 6" получает Быстрое Ожидание.' },
    part_of_plan: { name: 'Часть плана', category: 'Стратегия', text: 'Взаимодействие дальностью 12" на врага в ЛВ: встречная проверка Интеллекта — успех: недобровольное Передвижение цели (без падения); супер-успех: цель также Оглушена.' },
    outstanding_mind: { name: 'Выдающийся ум', category: 'Стратегия', text: 'Проверка Интеллекта с супер-успехом — быстрое действие Взаимодействия или статус Быстрое Ожидание.' },
    cunning: { name: 'Коварный', category: 'Стратегия', text: 'При ничьей в броске инициативы — вы выигрываете (если у обоих есть Коварный — всё ещё ничья).' },

    command_tone: { name: 'Командный тон', category: 'Вдохновение', text: 'Только лидер отряда или персонаж с Героическим броском. Взаимодействие дальностью 6" на дружественного: проверка Воли(4) — успех: цель немедленно действует; супер-успех: применивший получает быстрое действие.' },
    inspiring_presence: { name: 'Вдохновляющее присутствие', category: 'Вдохновение', text: 'Супер-успех на стандартном действии этого персонажа — отряд получает 1 Очко Мотивации.' },
    heroic_example: { name: 'Героический пример', category: 'Вдохновение', text: 'Первое за бой использование Последнего Шанса этим персонажем — дружественный в 12" снимает Жетон Активации.' },
    grim_resolve: { name: 'Мрачная решимость', category: 'Вдохновение', text: 'В начале активации можно потерять до 2 Здоровья и получить до 2 быстрых действий (нельзя при Здоровье ≤2).' },

    ruthless: { name: 'Безжалостный', category: 'Беспощадность', text: 'Против персонажа с Жетоном активации оружие получает Надёжное(1) и Бронебойность(1).' },
    duelist: { name: 'Дуэлянт', category: 'Беспощадность', text: 'Оружие персонажа получает Блок(1).' },
    counterattack: { name: 'Контрудар', category: 'Беспощадность', text: 'Враг заявил этого персонажа целью Сражения (до проверки) — персонаж получает Быстрое Ожидание(Сражение).' },
    finishing_blow: { name: 'Добивающий удар', category: 'Беспощадность', text: 'Атака Сражения ввела врага в Тяжёлое ранение — проверка Сражения(4): успех — враг удалён с поля, Выбывший из боя.' },

    sniper_stance: { name: 'Снайперская позиция', category: 'Точность', text: 'На террейне своего размера или больше — стрелковое оружие получает Меткое и Надёжное(1).' },
    relentless_barrage: { name: 'Неистовый залп', category: 'Точность', text: 'При Стрельбе с Очередь(X) — можно увеличить X на 2, тогда оружие получает жетон Разряженного оружия(1).' },
    quick_reload: { name: 'Быстрая перезарядка', category: 'Точность', text: 'После действия Рывка — быстрое действие Взаимодействия только на перезарядку оружия.' },
    steady_stance: { name: 'Устойчивое положение', category: 'Точность', text: 'Не двигался и не делал Рывок в активацию — в конце получает Ожидание(Стрельба).' },

    extra_limb: { name: 'Ещё одна конечность', category: 'Больше чем человек', text: 'Персонаж получает ещё одну руку и 1 слот инвентаря. Можно взять несколько раз.' },
    winged: { name: 'Крылатый', category: 'Больше чем человек', text: 'Персонаж получает тип передвижения Полёт.' },
    tough_hide: { name: 'Крепкий покров', category: 'Больше чем человек', text: 'Персонаж получает Броню 2.' },
    builtin_melee: { name: 'Встроенное оружие ближнего боя', category: 'Больше чем человек', text: 'Выберите профиль: Встроенное тяжёлое (дальн.1, урон3, Бронебойность(2)) либо Встроенное длинное (дальн.3, урон2, Бронебойность(1), Пролом(4+)). Не занимает слотов. В начале активации — быстрое действие Сражения. Можно взять несколько раз (2-й раз — вместо доп. быстрого действия оружие получает Лёгкое).' },
    builtin_ranged: { name: 'Встроенное стрелковое оружие', category: 'Больше чем человек', text: 'Выберите профиль: Огненное дыхание (дальн.10, урон3, Фокусированный огонь(0–6)) либо Сфокусированный выстрел (дальн.18, урон2, Бронебойность(5)). Занимает 1 слот, не занимает рук. Можно взять несколько раз (2-й раз — оружие получает Лёгкое).' },

    telepathy: { name: 'Телепатия', category: 'Духовные силы', text: 'Стоимость 2 Очка Мотивации. Взаимодействие дальностью 12 на любого. На врага — встречная проверка Интеллекта: успех — недобровольное стандартное действие (враг считает свой отряд враждебным, ваш — дружественным); супер-успех(1) — также Оглушение; супер-успех(2+) — Оглушение и Замедление. На союзника — проверка Интеллекта(3): успех — стандартное действие; супер-успех — также Быстрое Ожидание.' },
    healing_wave: { name: 'Целебная волна', category: 'Духовные силы', text: 'Стоимость 1 Очко Мотивации. Взаимодействие дальностью 8 на союзника/нейтрала: проверка Интеллекта(3) — успех восстанавливает 2 Здоровья; супер-успех(X) — ещё X Здоровья.' },
    invisibility_power: { name: 'Невидимость', category: 'Духовные силы', text: 'Стоимость 1 Очко Мотивации. Взаимодействие дальностью 6 на союзника/нейтрала: проверка Интеллекта(3) — успех даёт статус Скрытый; супер-успех — применивший тоже становится Скрытым.' },
    time_acceleration: { name: 'Ускорение времени', category: 'Духовные силы', text: 'Стоимость 2 Очка Мотивации. Взаимодействие дальностью 6 на союзника: проверка Интеллекта(5) — успех снимает Жетон Активации; супер-успех — также быстрое действие Перемещения.' },
    time_stop: { name: 'Остановка времени', category: 'Духовные силы', text: 'Стоимость 2 Очка Мотивации. Взаимодействие дальностью 6 на врага: встречная проверка Интеллекта — успех даёт Оглушение и Замедление; супер-успех — Жетон Активации и Замедление вместо обычного эффекта.' },

    discharge: { name: 'Разряд', category: 'Манифестация', text: 'Стоимость 1–2 Очка Мотивации. Действие Стрельбы профилем: дальн.12, урон4, Пролом(5+); при трате 2 Очков — Пролом(6+) и Меткое.' },
    destructive_pulse: { name: 'Разрушительный импульс', category: 'Манифестация', text: 'Стоимость 1 Очко Мотивации. Взаимодействие на себя: проверка Интеллекта(4). Все в 3" проходят проверку Силы(6) — провал: 3 урона и недобровольный Бросок на 1" (дальше при супер-успехе(X) на +X).' },
    terrain_control: { name: 'Контроль ландшафта', category: 'Манифестация', text: 'Стоимость 1 Очко Мотивации. Взаимодействие дальностью 6 на точку в ЛВ: проверка Интеллекта(4) — успех размещает элемент террейна размера 2; супер-успех — два таких, либо один размера 3.' },
    gate: { name: 'Врата', category: 'Манифестация', text: 'Стоимость 1–3 Очка Мотивации. Взаимодействие дальностью 6 на любого. На врага — встречная проверка Интеллекта: успех — недобровольная Телепортация (4" за потраченное Очко). На союзника — проверка Интеллекта(4): успех — Телепортация (4" за Очко).' },
    shield_power: { name: 'Щит', category: 'Манифестация', text: 'Стоимость 2 Очка Мотивации. Взаимодействие дальностью 6 на союзника: проверка Интеллекта(4) — успех даёт статус Защищён; супер-успех — Защищён ещё на одного союзника в 6".' }
  },

  weapons: {
    pistol: { name: 'Пистолет', kind: 'ranged', range: 12, damage: 2, traits: ['Бронебойность(1)', 'Лёгкое'], cost: 1 },
    smg: { name: 'Пистолет-пулемёт', kind: 'ranged', range: 12, damage: 2, traits: ['Очередь(3)', 'Лёгкое'], cost: 1 },
    revolver: { name: 'Револьвер', kind: 'ranged', range: 12, damage: 3, traits: ['Бронебойность(1)', 'Перезарядка(1)'], cost: 1 },
    sawnoff: { name: 'Обрез', kind: 'ranged', range: 8, damage: 2, traits: ['Меткое', 'Фокусированный огонь(0–6)', 'Перезарядка(1)'], cost: 1 },
    assault: { name: 'Автомат', kind: 'ranged', range: 18, damage: 2, traits: ['Пролом(4+)', 'Очередь(4)', 'Двуручное'], twohanded: true, cost: 2 },
    shotgun: { name: 'Дробовик', kind: 'ranged', range: 12, damage: 2, traits: ['Меткое', 'Фокусированный огонь(0–6)', 'Двуручное'], twohanded: true, cost: 2 },
    line_rifle: { name: 'Линейная винтовка', kind: 'ranged', range: 24, damage: 3, traits: ['Бронебойность(1)', 'Двуручное'], twohanded: true, cost: 2 },
    sniper_rifle: { name: 'Снайперская винтовка', kind: 'ranged', range: 36, damage: 4, traits: ['Бронебойность(2)', 'Перезарядка(1)', 'Двуручное'], twohanded: true, cost: 3 },
    mg: { name: 'Пулемёт', kind: 'ranged', range: 24, damage: 3, traits: ['Бронебойность(1)', 'Пролом(4+)', 'Очередь(5)', 'Требует Силу(д8)', 'Двуручное'], twohanded: true, reqStrength: 'd8', cost: 3 },
    am_rifle: { name: 'Антиматериальная винтовка', kind: 'ranged', range: 24, damage: 4, traits: ['Пролом(6+)', 'Перезарядка(1)', 'Двуручное'], twohanded: true, cost: 3, uniqueMax: 1, faction: 'organization' },

    brass: { name: 'Кастет', kind: 'melee', range: 1, damage: 2, traits: ['Лёгкое'], cost: 1 },
    knife: { name: 'Нож', kind: 'melee', range: 1, damage: 1, traits: ['Бронебойность(1)', 'Надёжное(1)', 'Лёгкое'], cost: 1 },
    sword: { name: 'Меч', kind: 'melee', range: 1, damage: 3, traits: ['Бронебойность(1)', 'Блок(1)'], cost: 2 },
    axe: { name: 'Топор', kind: 'melee', range: 1, damage: 3, traits: ['Бронебойность(2)'], cost: 2 },
    mace: { name: 'Булава', kind: 'melee', range: 1, damage: 3, traits: ['Пролом(5+)'], cost: 2 },
    polearm: { name: 'Древковое оружие', kind: 'melee', range: 2, damage: 3, traits: ['Бронебойность(1)'], cost: 2 },
    hammer2h: { name: 'Двуручный молот', kind: 'melee', range: 1, damage: 4, traits: ['Пролом(5+)', 'Отталкивание(2)', 'Требует Силу(д8)', 'Двуручное'], twohanded: true, reqStrength: 'd8', cost: 3 },
    polearm2h: { name: 'Двуручное древковое оружие', kind: 'melee', range: 2, damage: 4, traits: ['Бронебойность(2)', 'Требует Силу(д8)', 'Двуручное'], twohanded: true, reqStrength: 'd8', cost: 3 },
    sword2h: { name: 'Двуручный меч', kind: 'melee', range: 1, damage: 4, traits: ['Бронебойность(2)', 'Блок(1)', 'Требует Силу(д8)', 'Двуручное'], twohanded: true, reqStrength: 'd8', cost: 3 },

    giant_sword: { name: 'Гигантский меч', kind: 'melee', range: 2, damage: 4, traits: ['Бронебойность(3)', 'Блок(1)', 'Требует Силу(д10)', 'Двуручное'], twohanded: true, reqStrength: 'd10', cost: 3, uniqueMax: 2, faction: 'heavy' },
    rotary_cannon: { name: 'Вращающаяся автоматическая пушка', kind: 'ranged', range: 24, damage: 4, traits: ['Бронебойность(1)', 'Пролом(4+)', 'Очередь(6)', 'Требует Силу(д10)', 'Двуручное'], twohanded: true, reqStrength: 'd10', cost: 3, uniqueMax: 2, faction: 'heavy' }
  },

  equipment: {
    armor_light: { name: 'Защитное снаряжение', group: 'armor', armorBonus: 1, cost: 1, text: 'Броня +1. Не сочетается с другим защитным снаряжением.' },
    armor_heavy: { name: 'Тяжёлое защитное снаряжение', group: 'armor', armorBonus: 2, reqStrength: 'd8', cost: 2, text: 'Броня +2. Требует Силу(д8). Не сочетается с другим защитным снаряжением.' },
    armor_super: { name: 'Сверхтяжёлое защитное снаряжение', group: 'armor', armorBonus: 4, reqStrength: 'd12', cost: 3, text: 'Броня +4. Требует Силу(д12). Не сочетается с другим защитным снаряжением.' },
    backpack: { name: 'Рюкзак', grantsSlot: 1, cost: 1, text: 'Даёт 1 слот инвентаря.' },
    medkit: { name: 'Медицинский набор', slot: 1, cost: 1, text: 'Взаимодействие «Полевая медицина»: проверка Интеллекта(3) — восстанавливает до 2 Здоровья дружественной/нейтральной цели.' },
    mobility_gear: { name: 'Средства вертикального перемещения', cost: 1, text: 'При добровольном перемещении штрафы за террейн снижаются на 1".' },
    multitool: { name: 'Мультитул', cost: 1, text: 'Сложность проверок при Взаимодействии с предметами/объектами −1.' },
    targeting: { name: 'Система наведения', slot: 1, cost: 1, text: 'Взаимодействие на врага в ЛВ и 12": проверка Интеллекта(4) — цель получает Отмеченный.' },
    flight_gear: { name: 'Индивидуальное средство полёта', slot: 1, cost: 2, text: 'При добровольном перемещении всегда используется тип Полёт.' },
    camo: { name: 'Камуфлирующее покрытие', cost: 1, text: 'Взаимодействие на себя: проверка Интеллекта(4) — получает статус Скрытый.' },
    tac_shield: { name: 'Тактический щит', slot: 1, reqStrength: 'd8', cost: 2, text: 'Занимает руку. В начале активации персонаж получает статус Защищён.' },
    drone: { name: 'Тактический дрон', slot: 1, cost: 2, uniqueMax: 2, faction: 'organization', text: 'Взаимодействие на врага в 12": проверка Интеллекта(4) — цель теряет Скрытый и получает Отмеченный.' },
    taser: { name: 'Тазер', slot: 1, cost: 1, uniqueMax: 2, faction: 'organization', text: 'Особое действие Стрельбы на врага в 8" и ЛВ: при успехе цель получает Замедление и Оглушение.' },

    talisman: { name: 'Защитный талисман', armorBonusExtra: 1, cost: 2, uniqueMax: 3, faction: 'gifted', text: 'Броня +1 (складывается с обычным защитным снаряжением). Персонаж игнорирует спецправило Бронебойность у атак по нему.' },
    amplifier: { name: 'Амплификатор силы', slot: 1, cost: 1, uniqueMax: 2, faction: 'gifted', text: 'Занимает руку. Траты Очков Мотивации на Сверхъестественные силы не учитываются в лимит эффектов Мотивации за активацию; сложность проверок Интеллекта для Сверхъестественных сил −1.' },
    empowered_weapon: { name: 'Оружие, наделённое силой', cost: 2, uniqueMax: 3, faction: 'gifted', text: 'Улучшение для ОДНОГО оружия персонажа (привяжите вручную): урон +1; ближнего боя — Бронебойность(1); стрелковое — Меткое.' },

    unknown_armor: { name: 'Неизвестный тип брони', group: 'armor', armorBonus: 1, cost: 2, faction: 'heavy', text: 'Броня +1. Сложность проверок Брони никогда не выше 4.' }
  },

  glossary: {
    ap: { name: 'Бронебойность(X)', text: 'уменьшает Броню цели на X на время атаки.' },
    breach: { name: 'Пролом(X)', text: 'увеличивает сложность проверки Брони цели до X на время атаки.' },
    burst: { name: 'Очередь(X)', text: 'если есть 2+ действия — тратит их все, меняет Урон оружия на X.' },
    reliable: { name: 'Надёжное(X)', text: 'при промахе всё равно наносит X урона.' },
    block: { name: 'Блок(X)', text: 'увеличивает Защиту персонажа в ближнем бою на X.' },
    reload: { name: 'Перезарядка(X)', text: 'после Стрельбы — жетон Разряженного оружия(X); снимается X действиями Взаимодействия, пока висит — стрелять нельзя.' },
    twohanded: { name: 'Двуручное', text: 'занимает 2 слота инвентаря и обе руки.' },
    focusfire: { name: 'Фокусированный огонь(X–Y)', text: 'в пределах дистанции оружие получает Урон +1 и Надёжное(1).' },
    accurate: { name: 'Меткое', text: 'игнорирует модификатор к Защите (не к Броне) от укрытия.' },
    light: { name: 'Лёгкое', text: 'снижает штраф от Быстрого действия на 1.' },
    reqstat: { name: 'Требует Параметр(X)', text: 'нельзя экипировать при значении параметра меньше X.' },
    knockback: { name: 'Отталкивание(X)', text: 'при попадании после урона можно совершить Бросок цели на X.' }
  }
};

/* ============================================================ СИДИНГ ==== */

function seedAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  seedWeapons(ss);
  seedEquipment(ss);
  seedAbilities(ss);
  seedGlossary(ss);
  seedArchetypesAndRoles(ss);
  seedSpecialRules(ss);
  ensureSquadsSheet(ss);
  SpreadsheetApp.getUi().alert('Готово! Вкладки с данными созданы/проверены. Изменяйте значения прямо в ячейках — сайт подхватит их при следующем открытии.');
}

function getOrCreateSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function sheetHasDataRows_(sh) {
  return sh.getLastRow() > 1;
}

function seedWeapons(ss) {
  var headers = ['id', 'name', 'kind', 'range', 'damage', 'traits (через запятую)', 'cost', 'twohanded (TRUE/FALSE)', 'reqStrength (d8/d10/d12, необязательно)', 'uniqueMax (необязательно)', 'faction (organization/gifted/heavy, необязательно)'];
  var sh = getOrCreateSheet_(ss, SHEET_WEAPONS, headers);
  if (sheetHasDataRows_(sh)) return;
  var rows = [];
  Object.keys(SEED_DATA.weapons).forEach(function (id) {
    var w = SEED_DATA.weapons[id];
    rows.push([id, w.name, w.kind, w.range, w.damage, (w.traits || []).join(', '), w.cost, !!w.twohanded, w.reqStrength || '', w.uniqueMax || '', w.faction || '']);
  });
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function seedEquipment(ss) {
  var headers = ['id', 'name', 'group (armor — для взаимоисключающей брони, иначе пусто)', 'armorBonus', 'armorBonusExtra (складывается с бронёй)', 'grantsSlot', 'slot (сколько слотов инвентаря занимает)', 'reqStrength (необязательно)', 'cost', 'uniqueMax (необязательно)', 'faction (необязательно)', 'text (описание)'];
  var sh = getOrCreateSheet_(ss, SHEET_EQUIPMENT, headers);
  if (sheetHasDataRows_(sh)) return;
  var rows = [];
  Object.keys(SEED_DATA.equipment).forEach(function (id) {
    var e = SEED_DATA.equipment[id];
    rows.push([id, e.name, e.group || '', e.armorBonus || '', e.armorBonusExtra || '', e.grantsSlot || '', e.slot || '', e.reqStrength || '', e.cost, e.uniqueMax || '', e.faction || '', e.text || '']);
  });
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function seedAbilities(ss) {
  var headers = ['id', 'name', 'category', 'text (описание)'];
  var sh = getOrCreateSheet_(ss, SHEET_ABILITIES, headers);
  if (sheetHasDataRows_(sh)) return;
  var rows = [];
  Object.keys(SEED_DATA.abilities).forEach(function (id) {
    var a = SEED_DATA.abilities[id];
    rows.push([id, a.name, a.category, a.text]);
  });
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function seedGlossary(ss) {
  var headers = ['id', 'name', 'text (описание)'];
  var sh = getOrCreateSheet_(ss, SHEET_GLOSSARY, headers);
  if (sheetHasDataRows_(sh)) return;
  var rows = [];
  Object.keys(SEED_DATA.glossary).forEach(function (id) {
    var g = SEED_DATA.glossary[id];
    rows.push([id, g.name, g.text]);
  });
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function seedArchetypesAndRoles(ss) {
  var archHeaders = ['id', 'name', 'blurb (описание архетипа)', 'archetypeCategories (категории способностей через запятую)', 'sizeBonusRule_size (необязательно)', 'sizeBonusRule_armor', 'sizeBonusRule_health'];
  var archSh = getOrCreateSheet_(ss, SHEET_ARCHETYPES, archHeaders);
  var roleHeaders = ['archetypeId', 'roleId', 'name', 'tier (senior/novice)', 'maxCount', 'sizeMin', 'sizeMax', 'defaultSize', 'cost', 'level', 'statIncreases', 'archetypeAbilitySlots', 'anyAbilitySlots', 'builtinArmor', 'builtinHealth', 'mightBaseline (d8 и т.п., необязательно)', 'speedCapped (TRUE/FALSE)', 'archetypeCategories override (через запятую, необязательно)', 'note (необязательно)'];
  var roleSh = getOrCreateSheet_(ss, SHEET_ROLES, roleHeaders);

  if (!sheetHasDataRows_(archSh)) {
    var archRows = [];
    Object.keys(SEED_DATA.archetypes).forEach(function (aid) {
      var a = SEED_DATA.archetypes[aid];
      var sb = a.sizeBonusRule || {};
      archRows.push([aid, a.name, a.blurb, (a.archetypeCategories || []).join(', '), sb.size || '', sb.armor || '', sb.health || '']);
    });
    archSh.getRange(2, 1, archRows.length, archHeaders.length).setValues(archRows);
  }

  if (!sheetHasDataRows_(roleSh)) {
    var roleRows = [];
    Object.keys(SEED_DATA.archetypes).forEach(function (aid) {
      var a = SEED_DATA.archetypes[aid];
      Object.keys(a.roles).forEach(function (rid) {
        var r = a.roles[rid];
        roleRows.push([
          aid, rid, r.name, r.tier, r.maxCount, r.sizeMin, r.sizeMax, r.defaultSize, r.cost, r.level, r.statIncreases,
          r.archetypeAbilitySlots, r.anyAbilitySlots, r.builtinArmor || '', r.builtinHealth || '', r.mightBaseline || '',
          !!r.speedCapped, (r.archetypeCategories || []).join(', '), r.note || ''
        ]);
      });
    });
    roleSh.getRange(2, 1, roleRows.length, roleHeaders.length).setValues(roleRows);
  }
}

function seedSpecialRules(ss) {
  var headers = ['archetypeId', 'name', 'text'];
  var sh = getOrCreateSheet_(ss, SHEET_SPECIAL_RULES, headers);
  if (sheetHasDataRows_(sh)) return;
  var rows = [];
  Object.keys(SEED_DATA.archetypes).forEach(function (aid) {
    var a = SEED_DATA.archetypes[aid];
    (a.specialRules || []).forEach(function (r) {
      rows.push([aid, r.name, r.text]);
    });
  });
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function ensureSquadsSheet(ss) {
  var headers = ['id', 'ownerName', 'accessCode', 'squadName', 'archetype', 'budget', 'unitsJson', 'updatedAt'];
  getOrCreateSheet_(ss, SHEET_SQUADS, headers);
}

/* ==================================================== ЧТЕНИЕ ДАННЫХ ==== */

function sheetRows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return values
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function num_(v) { return (v === '' || v === null || v === undefined) ? undefined : Number(v); }
function str_(v) { return (v === '' || v === null || v === undefined) ? undefined : String(v).trim(); }
function bool_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1 || v === '1' || v === 'да'; }
function list_(v) {
  var s = String(v || '').trim();
  if (!s) return [];
  return s.split(',').map(function (x) { return x.trim(); }).filter(Boolean);
}
function compact_(obj) {
  var out = {};
  Object.keys(obj).forEach(function (k) { if (obj[k] !== undefined && obj[k] !== '') out[k] = obj[k]; });
  return out;
}

function buildGameData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var weapons = {};
  sheetRows_(ss.getSheetByName(SHEET_WEAPONS)).forEach(function (row) {
    var headers = Object.keys(row);
    var id = str_(row['id']);
    if (!id) return;
    weapons[id] = compact_({
      name: str_(row['name']),
      kind: str_(row['kind']),
      range: num_(row['range']),
      damage: num_(row['damage']),
      traits: list_(row['traits (через запятую)']),
      cost: num_(row['cost']),
      twohanded: bool_(row['twohanded (TRUE/FALSE)']) || undefined,
      reqStrength: str_(row['reqStrength (d8/d10/d12, необязательно)']),
      uniqueMax: num_(row['uniqueMax (необязательно)']),
      faction: str_(row['faction (organization/gifted/heavy, необязательно)'])
    });
    if (weapons[id].traits && weapons[id].traits.length === 0) delete weapons[id].traits;
  });

  var equipment = {};
  sheetRows_(ss.getSheetByName(SHEET_EQUIPMENT)).forEach(function (row) {
    var id = str_(row['id']);
    if (!id) return;
    equipment[id] = compact_({
      name: str_(row['name']),
      group: str_(row['group (armor — для взаимоисключающей брони, иначе пусто)']),
      armorBonus: num_(row['armorBonus']),
      armorBonusExtra: num_(row['armorBonusExtra (складывается с бронёй)']),
      grantsSlot: num_(row['grantsSlot']),
      slot: num_(row['slot (сколько слотов инвентаря занимает)']),
      reqStrength: str_(row['reqStrength (необязательно)']),
      cost: num_(row['cost']),
      uniqueMax: num_(row['uniqueMax (необязательно)']),
      faction: str_(row['faction (необязательно)']),
      text: str_(row['text (описание)'])
    });
  });

  var abilities = {};
  sheetRows_(ss.getSheetByName(SHEET_ABILITIES)).forEach(function (row) {
    var id = str_(row['id']);
    if (!id) return;
    abilities[id] = { name: str_(row['name']), category: str_(row['category']), text: str_(row['text (описание)']) };
  });

  var glossary = {};
  sheetRows_(ss.getSheetByName(SHEET_GLOSSARY)).forEach(function (row) {
    var id = str_(row['id']);
    if (!id) return;
    glossary[id] = { name: str_(row['name']), text: str_(row['text (описание)']) };
  });

  var specialRulesByArch = {};
  sheetRows_(ss.getSheetByName(SHEET_SPECIAL_RULES)).forEach(function (row) {
    var aid = str_(row['archetypeId']);
    if (!aid) return;
    specialRulesByArch[aid] = specialRulesByArch[aid] || [];
    specialRulesByArch[aid].push({ name: str_(row['name']), text: str_(row['text']) });
  });

  var rolesByArch = {};
  sheetRows_(ss.getSheetByName(SHEET_ROLES)).forEach(function (row) {
    var aid = str_(row['archetypeId']);
    var rid = str_(row['roleId']);
    if (!aid || !rid) return;
    rolesByArch[aid] = rolesByArch[aid] || {};
    rolesByArch[aid][rid] = compact_({
      id: rid,
      name: str_(row['name']),
      tier: str_(row['tier (senior/novice)']),
      maxCount: num_(row['maxCount']),
      sizeMin: num_(row['sizeMin']),
      sizeMax: num_(row['sizeMax']),
      defaultSize: num_(row['defaultSize']),
      cost: num_(row['cost']),
      level: num_(row['level']),
      statIncreases: num_(row['statIncreases']),
      archetypeAbilitySlots: num_(row['archetypeAbilitySlots']),
      anyAbilitySlots: num_(row['anyAbilitySlots']),
      builtinArmor: num_(row['builtinArmor']),
      builtinHealth: num_(row['builtinHealth']),
      mightBaseline: str_(row['mightBaseline (d8 и т.п., необязательно)']),
      speedCapped: bool_(row['speedCapped (TRUE/FALSE)']) || undefined,
      archetypeCategories: list_(row['archetypeCategories override (через запятую, необязательно)']),
      note: str_(row['note (необязательно)'])
    });
    if (rolesByArch[aid][rid].archetypeCategories && rolesByArch[aid][rid].archetypeCategories.length === 0) delete rolesByArch[aid][rid].archetypeCategories;
  });

  var archetypes = {};
  sheetRows_(ss.getSheetByName(SHEET_ARCHETYPES)).forEach(function (row) {
    var aid = str_(row['id']);
    if (!aid) return;
    var sizeBonusRule;
    var sbSize = num_(row['sizeBonusRule_size (необязательно)']);
    if (sbSize) sizeBonusRule = { size: sbSize, armor: num_(row['sizeBonusRule_armor']) || 0, health: num_(row['sizeBonusRule_health']) || 0 };
    archetypes[aid] = compact_({
      id: aid,
      name: str_(row['name']),
      blurb: str_(row['blurb (описание архетипа)']),
      archetypeCategories: list_(row['archetypeCategories (категории способностей через запятую)']),
      specialRules: specialRulesByArch[aid] || [],
      sizeBonusRule: sizeBonusRule,
      roles: rolesByArch[aid] || {}
    });
  });

  return { archetypes: archetypes, abilities: abilities, weapons: weapons, equipment: equipment, glossary: glossary };
}

/* ============================================================ ОТРЯДЫ ==== */

function normName_(s) { return String(s || '').trim().toLowerCase(); }

function listSquads_(ownerName, accessCode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_SQUADS);
  var rows = sheetRows_(sh);
  return rows
    .filter(function (r) { return normName_(r.ownerName) === normName_(ownerName) && String(r.accessCode) === String(accessCode); })
    .map(function (r) {
      var units = [];
      try { units = JSON.parse(r.unitsJson || '[]'); } catch (e) { units = []; }
      return { id: r.id, squadName: r.squadName, archetype: r.archetype, budget: Number(r.budget), units: units, updatedAt: r.updatedAt };
    });
}

function saveSquad_(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_SQUADS);
  var ownerName = str_(payload.ownerName);
  var accessCode = str_(payload.accessCode);
  if (!ownerName || !accessCode) throw new Error('Укажите имя и код доступа.');

  var id = str_(payload.id);
  var lastRow = sh.getLastRow();
  var foundRowIndex = -1;
  if (id && lastRow > 1) {
    var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) { foundRowIndex = i + 2; break; }
    }
  }

  var now = new Date().toISOString();

  if (foundRowIndex > -1) {
    var existing = sh.getRange(foundRowIndex, 1, 1, 8).getValues()[0];
    if (normName_(existing[1]) !== normName_(ownerName) || String(existing[2]) !== accessCode) {
      throw new Error('Отряд с таким id принадлежит другому имени/коду доступа.');
    }
    sh.getRange(foundRowIndex, 4, 1, 5).setValues([[
      payload.squadName || '', payload.archetype || '', payload.budget || 0, JSON.stringify(payload.units || []), now
    ]]);
    return { id: id, updatedAt: now };
  } else {
    var newId = Utilities.getUuid();
    sh.appendRow([newId, ownerName, accessCode, payload.squadName || '', payload.archetype || '', payload.budget || 0, JSON.stringify(payload.units || []), now]);
    return { id: newId, updatedAt: now };
  }
}

function deleteSquad_(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_SQUADS);
  var id = str_(payload.id);
  var ownerName = str_(payload.ownerName);
  var accessCode = str_(payload.accessCode);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true };
  var values = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      if (normName_(values[i][1]) !== normName_(ownerName) || String(values[i][2]) !== accessCode) {
        throw new Error('Нет прав удалить этот отряд.');
      }
      sh.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: true };
}

/* ========================================================== HTTP API ==== */

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ensureSeeded_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_WEAPONS);
  if (!sh || sh.getLastRow() < 2) seedAll();
}

function doGet(e) {
  try {
    ensureSeeded_();
    var action = (e && e.parameter && e.parameter.action) || 'data';
    if (action === 'ping') return jsonOut_({ ok: true, ping: 'pong' });
    return jsonOut_({ ok: true, data: buildGameData() });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents || '{}');
    var action = payload.action;
    if (action === 'listSquads') {
      return jsonOut_({ ok: true, squads: listSquads_(payload.ownerName, payload.accessCode) });
    }
    if (action === 'saveSquad') {
      return jsonOut_({ ok: true, result: saveSquad_(payload) });
    }
    if (action === 'deleteSquad') {
      return jsonOut_({ ok: true, result: deleteSquad_(payload) });
    }
    return jsonOut_({ ok: false, error: 'Неизвестное действие: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
