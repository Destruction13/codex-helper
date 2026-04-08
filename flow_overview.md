# Codex Helper flow overview

Этот файл показывает текущую логику полного цикла `reg` в `codex_helper_app.py`.

```mermaid
flowchart TD
    A[Старт кнопки reg] --> B[Поднять reg_automation_active и reg_in_progress]
    B --> C[Открыть ChatGPT в MS Edge InPrivate]
    C --> D[Подождать старт страницы и нажать Зарегистрироваться бесплатно]
    D --> E[Вставить текущую почту и отправить]
    E --> F[Watcher регистрации ChatGPT]

    F -->|Найден шаг Пароль| G[Вставить CHATGPT_ACCOUNT_PASSWORD]
    G --> H[Нажать Enter и ждать код]

    F -->|Найден шаг Проверьте свою почту| H

    H --> I[Ждать код из Zoho IMAP]
    I -->|Код не пришёл 30 секунд| J[Нажать Отправить электронное письмо повторно в MS Edge]
    J --> K[Вернуть фокус в поле кода]
    K --> I

    I -->|Лимит resend исчерпан| L[Закрыть MS Edge]
    L --> M[+1 к alias]
    M --> N[Следующая итерация reg]

    I -->|Код пришёл| O[Вставить код и нажать Enter]
    O --> P[Post-code watcher ChatGPT]

    P -->|Неверный код| Q[Найти поле Код в Edge, очистить его]
    Q --> R[Ждать новый код]
    R -->|Новый код пришёл| O
    R -->|Код не пришёл 30 секунд| J

    P -->|Код принят| S[Шаг profile setup]
    S --> T[Вставить имя]
    T --> U[Определить формат age/date step]
    U -->|ГГГГ| V[Вставить день, месяц, год]
    U -->|2026| W[Два Tab и вставить год]
    U -->|Иное| X[Вставить возраст]

    V --> Y[Нажать Завершить создание учетной записи или Enter]
    W --> Y
    X --> Y

    Y --> Z[Подождать 2 секунды]
    Z --> AA[Закрыть окно OpenAI в MS Edge]
    AA --> AB[Старт OmniRoute]

    AB --> AC[Поднять SSH туннель если нужно]
    AC --> AD{Reuse OmniRoute вкладки?}
    AD -->|Да| AE[Активировать существующую вкладку Yandex]
    AD -->|Нет| AF[Открыть новую вкладку OmniRoute]
    AF --> AG[Дождаться готовности страницы OpenAI Codex / Connections]
    AE --> AH[Поиск кнопки Add]
    AG --> AH

    AH -->|Add найден| AI[Нажать Add]
    AI --> AJ[Подождать 5 секунд]
    AJ --> AK[Вставить текущую почту и нажать Enter]
    AK --> AL[Watcher шага OmniRoute]

    AL -->|Шаг Пароль| AM[Вставить CHATGPT_ACCOUNT_PASSWORD]
    AM --> AN[Нажать Enter и ждать код]

    AL -->|Сразу шаг Проверьте свою почту| AN

    AN --> AO[Ждать код из Zoho IMAP]
    AO -->|Код не пришёл 30 секунд| AP[Нажать Отправить электронное письмо повторно в Yandex]
    AP --> AQ[Вернуть фокус в поле кода]
    AQ --> AO

    AO -->|Лимит resend исчерпан| AR[Закрыть 2 вкладки Yandex]
    AR --> AS[+1 к alias]
    AS --> N

    AO -->|Код пришёл| AT[Вставить код и нажать Enter]
    AT --> AU[Post-code watcher OmniRoute]

    AU -->|Неверный код| AV[Найти поле Код в Yandex, очистить его]
    AV --> AW[Ждать новый код]
    AW -->|Новый код пришёл| AT
    AW -->|Код не пришёл 30 секунд| AP

    AU -->|Требуется номер телефона или возраст| AR

    AU -->|Найден финальный Продолжить| AX[Нажать Продолжить]
    AX --> AY[Успешное завершение OmniRoute]
    AY --> AZ[+1 к alias]
    AZ --> BA{reg_automation_active?}
    BA -->|Да| N
    BA -->|Нет| BB[Цикл завершён]
```

## Основные точки отказа

- `ChatGPT` шаг resend и invalid-code retry работают только в `MS Edge`.
- `OmniRoute` шаг resend и invalid-code retry работают только в `Yandex Browser`.
- После reject-сценария `OmniRoute` вкладка reuse сбрасывается, и следующая
  итерация открывает `OmniRoute` заново как первую.
- Inspector на `стрелке вниз` не меняет поток, а только собирает debug-дамп
  активного окна.

## Где смотреть код

- Основной файл: `codex_helper_app.py`
- Логи: `codex_helper.log`
- Inspector hotkey: `стрелка вниз`
