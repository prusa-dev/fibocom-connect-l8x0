# Fibocom L8x0 Connect для Windows

![](./screenshot/screen01.png)

## Запуск

Все скрипты **_должны_** запускаться с правами администратора

- `connect.cmd`: Подключение и мониторинг
- `monitor.cmd`: Мониторинг соединения без подключения

## Настройка

#### APN

Отредактируйте `scripts/main.ps1` чтобы настроить APN, APN_USER and APN_PASS своего оператора

#### Предпочитаемые бенды

Найдите `AT+XACT=` в файле `scripts/main.ps1` и отредактируйте в соостветсвии со своими предпочтениями

Пример:

- UMTS+LTE all bands, LTE preferred: AT+XACT=4,2,,0
- LTE all bands: AT+XACT=2,,,0
- LTE 3 and 7 bands: AT+XACT=2,,,103,107

#### Установка своих DNS серверов

Отредактируйте `scripts/main.ps1` чтобы настроить свои DNS сервера: DNS_OVERRIDE
