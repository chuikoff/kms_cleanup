# 🧹 KMS Cleanup Tool

## 🇷🇺 Русская версия

Скрипт PowerShell для удаления последствий KMS-активаторов и восстановления системных настроек Windows.

---

### 🔧 Возможности

- Аудит системы перед изменениями
- Интерактивное подтверждение действий
- Удаление:
  - задач планировщика
  - сервисов
  - автозагрузки (Run)
  - исключений Windows Defender
- Проверка и восстановление Defender
- Сброс KMS-активации
- Логирование всех действий

---

### ⚙️ Использование

#### 1. Разрешить выполнение скриптов (временно)

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
Подтвердите выбор:

Y

2. Проверка (рекомендуется сначала)
.\kms_cleanup.ps1 -DryRun

Показывает, что будет сделано, без внесения изменений.

3. Интерактивный режим
.\kms_cleanup.ps1

Скрипт будет спрашивать подтверждение на каждое действие.

4. Автоматический режим
.\kms_cleanup.ps1 -AutoApprove

Все действия выполняются автоматически.
🛡️ Особенности
Defender не изменяется, если уже работает корректно
Безопасная обработка ошибок PowerShell
Защита от null/пустых значений
Учитываются ограничения:
Tamper Protection
Group Policy
⚠️ Ограничения
Не является полноценным антивирусом
Не удаляет сложные механизмы закрепления (WMI, rootkits)
Может не удалить вручную установленные компоненты активаторов
🔄 После выполнения

Рекомендуется:

Перезагрузить компьютер
Активировать Windows легальным ключом

English Version
📌 Description

PowerShell script designed to remove traces of KMS activators (such as AAct, KMSAuto, etc.) and restore Windows security settings.

Workflow:

audit → confirmation → remediation

🔧 Features
Pre-change system audit
Interactive confirmation for each action
Removes:
Scheduled tasks
Services
Run entries
Windows Defender exclusions
Checks Defender status
Restores Defender if disabled
Resets Windows activation (removes KMS)
Full logging to file
⚙️ Usage
1. Allow script execution (temporary)
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
2. Dry run (recommended)
.\kms_cleanup.ps1 -DryRun

Shows planned actions without making changes.

3. Interactive mode
.\kms_cleanup.ps1

Prompts for confirmation before each action.

4. Automatic mode
.\kms_cleanup.ps1 -AutoApprove

Runs without prompts.

🛡️ Notes
Defender is not modified if already running correctly
Handles PowerShell edge cases safely
Prevents null/empty argument errors
Some actions may be limited by:
Tamper Protection
Group Policy
⚠️ Limitations
Not a full antivirus solution
Does not detect advanced persistence techniques (e.g. WMI, rootkits)
May not remove manually installed activator components
🔄 After running

Recommended:

Reboot system
Activate Windows using a valid license
