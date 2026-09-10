import React, { useState } from 'react';
import { 
  Copy, 
  Check, 
  Download, 
  Zap, 
  FileSpreadsheet,
  Settings,
  ShieldCheck,
  CheckCircle2,
  Calendar,
  Terminal,
  Layers,
  FolderOpen,
  UserCheck,
  Play,
  Activity,
  Gauge
} from 'lucide-react';

export default function App() {
  const [copiedCode, setCopiedCode] = useState(false);
  const [copiedCmd, setCopiedCmd] = useState(false);
  const [activeTab, setActiveTab] = useState<'overview' | 'code' | 'guide'>('overview');
  const [mockPing, setMockPing] = useState({
    ms: 145,
    label: 'Низкая (Сервер свободен)',
    color: 'text-emerald-400',
    delay: 1.2,
    isChecking: false
  });

  const handleSimulatePing = () => {
    setMockPing(prev => ({ ...prev, isChecking: true }));
    setTimeout(() => {
      const ms = Math.floor(Math.random() * 120) + 95;
      setMockPing({
        ms,
        label: 'Низкая (Сервер свободен)',
        color: 'text-emerald-400',
        delay: 1.2,
        isChecking: false
      });
    }, 450);
  };

  const handleCopyCommand = () => {
    navigator.clipboard.writeText("pip install requests openpyxl\npython navigator_app.py");
    setCopiedCmd(true);
    setTimeout(() => setCopiedCmd(false), 2000);
  };

  const handleDownloadApp = () => {
    const a = document.createElement('a');
    a.href = '/navigator_app.py';
    a.download = 'navigator_app.py';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const handleDownloadExcel = () => {
    const a = document.createElement('a');
    a.href = '/list.xlsx';
    a.download = 'list.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const handleDownloadConfig = () => {
    const a = document.createElement('a');
    a.href = '/config.json';
    a.download = 'config.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const handleDownloadLog = () => {
    const a = document.createElement('a');
    a.href = '/results_log.txt';
    a.download = 'results_log.txt';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans selection:bg-blue-600 selection:text-white">
      {/* Top Header */}
      <header className="border-b border-slate-800 bg-slate-900/80 backdrop-blur sticky top-0 z-20 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-blue-600 flex items-center justify-center font-bold text-white shadow-lg shadow-blue-500/20">
            <Zap className="w-5 h-5 text-amber-300" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-base font-bold text-white tracking-tight">Навигатор ДО: Автоматизация (GUI)</h1>
              <span className="px-2 py-0.5 rounded-full text-[11px] font-medium bg-emerald-500/15 text-emerald-400 border border-emerald-500/30">
                Ctrl+C / Ctrl+V + ПКМ
              </span>
            </div>
            <p className="text-xs text-slate-400">
              Графический интерфейс, авто-вход, поддержка буфера обмена в RU/EN раскладках, чтение list.xlsx
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          <button
            onClick={handleDownloadExcel}
            className="px-3.5 py-1.5 rounded-lg border border-emerald-500/30 bg-emerald-500/10 hover:bg-emerald-500/20 text-emerald-300 text-xs font-medium flex items-center gap-1.5 transition-colors"
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" />
            Скачать list.xlsx
          </button>
          <button
            onClick={handleDownloadLog}
            className="px-3.5 py-1.5 rounded-lg border border-blue-500/30 bg-blue-500/10 hover:bg-blue-500/20 text-blue-300 text-xs font-medium flex items-center gap-1.5 transition-colors"
          >
            <Layers className="w-3.5 h-3.5 text-blue-400" />
            results_log.txt
          </button>
          <button
            onClick={handleDownloadConfig}
            className="px-3.5 py-1.5 rounded-lg border border-slate-700 hover:bg-slate-800 text-slate-300 text-xs font-medium flex items-center gap-1.5 transition-colors"
          >
            <Settings className="w-3.5 h-3.5 text-slate-400" />
            config.json
          </button>
          <button
            onClick={handleDownloadApp}
            className="px-4 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold flex items-center gap-1.5 transition-all shadow-md shadow-blue-600/20"
          >
            <Download className="w-3.5 h-3.5" />
            Скачать navigator_app.py (GUI)
          </button>
        </div>
      </header>

      {/* Navigation tabs */}
      <div className="border-b border-slate-800/80 bg-slate-900/40 px-6">
        <div className="max-w-6xl mx-auto flex items-center gap-4 text-xs font-medium">
          <button
            onClick={() => setActiveTab('overview')}
            className={`py-3 border-b-2 flex items-center gap-2 transition-colors ${
              activeTab === 'overview'
                ? 'border-blue-500 text-blue-400 font-semibold'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <Layers className="w-4 h-4" />
            Обзор интерфейса и возможностей
          </button>
          <button
            onClick={() => setActiveTab('guide')}
            className={`py-3 border-b-2 flex items-center gap-2 transition-colors ${
              activeTab === 'guide'
                ? 'border-blue-500 text-blue-400 font-semibold'
                : 'border-transparent text-slate-400 hover:text-slate-200'
            }`}
          >
            <Terminal className="w-4 h-4" />
            Инструкция запуска в 1 шаг
          </button>
        </div>
      </div>

      {/* Main Content Area */}
      <main className="flex-1 max-w-6xl w-full mx-auto p-6 space-y-6">
        {activeTab === 'overview' && (
          <div className="space-y-6">
            {/* 4 Feature Highlights */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4 flex flex-col justify-between space-y-3">
                <div className="flex items-start gap-3">
                  <div className="p-2 rounded-lg bg-blue-500/10 text-blue-400 border border-blue-500/20">
                    <UserCheck className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-sm font-semibold text-slate-200">Авто-сохранение</h2>
                    <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                      Логин, пароль, <b>название мероприятия</b> и <b>дата/время</b> сохраняются в <code>config.json</code>.
                    </p>
                  </div>
                </div>
                <div className="text-[11px] text-blue-400/90 bg-blue-950/30 p-2 rounded border border-blue-500/20">
                  Подсказки с примером ввода (ГГГГ-ММ-ДД)
                </div>
              </div>

              <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4 flex flex-col justify-between space-y-3">
                <div className="flex items-start gap-3">
                  <div className="p-2 rounded-lg bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-sm font-semibold text-slate-200">Таблица list.xlsx (3 столбца)</h2>
                    <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                      Столбцы: <b>ФИО</b>, <b>Дата рождения</b> и <b>Статус заявки</b>. Кнопка <b>«📂 Открыть list.xlsx»</b> сразу открывает файл в Excel.
                    </p>
                  </div>
                </div>
                <div className="text-[11px] text-emerald-400/90 bg-emerald-950/30 p-2 rounded border border-emerald-500/20 flex items-center gap-1.5">
                  <CheckCircle2 className="w-3.5 h-3.5" />
                  Авто-запись флага «Добавлен»
                </div>
              </div>

              <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4 flex flex-col justify-between space-y-3">
                <div className="flex items-start gap-3">
                  <div className="p-2 rounded-lg bg-amber-500/10 text-amber-400 border border-amber-500/20">
                    <ShieldCheck className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-sm font-semibold text-slate-200">Защита от дублирования</h2>
                    <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                      При повторном запуске ранее добавленные дети <b>игнорируются</b>. В лог выводится число пропущенных дубликатов.
                    </p>
                  </div>
                </div>
                <div className="text-[11px] text-amber-400/90 bg-amber-950/30 p-2 rounded border border-amber-500/20">
                  Проверка is_approved и дат тезок
                </div>
              </div>

              <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4 flex flex-col justify-between space-y-3">
                <div className="flex items-start gap-3">
                  <div className="p-2 rounded-lg bg-indigo-500/10 text-indigo-400 border border-indigo-500/20">
                    <Gauge className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-sm font-semibold text-slate-200">Защита от перегрузки</h2>
                    <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                      Строго <b>≤ 1 запрос в сек</b> (пауза 1.2 с). Адаптивное замедление при росте задержки и обработка 429/503.
                    </p>
                  </div>
                </div>
                <div className="text-[11px] text-indigo-400/90 bg-indigo-950/30 p-2 rounded border border-indigo-500/20 flex items-center gap-1.5">
                  <Activity className="w-3.5 h-3.5" />
                  Монитор пинга в реальном времени
                </div>
              </div>
            </div>

            {/* Interactive Mockup of the Desktop GUI */}
            <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-2xl">
              <div className="px-4 py-3 bg-slate-950 border-b border-slate-800 flex items-center justify-between text-xs text-slate-400">
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 rounded-full bg-rose-500/80" />
                  <div className="w-3 h-3 rounded-full bg-amber-500/80" />
                  <div className="w-3 h-3 rounded-full bg-emerald-500/80" />
                  <span className="ml-2 font-medium text-slate-300">Навигатор: Автоматизация (NavAdd & NavConfirm) — Графическое окно</span>
                </div>
                <span className="text-[11px] font-mono text-slate-500">Tkinter Desktop App</span>
              </div>

              <div className="p-6 bg-slate-900/70 space-y-5">
                {/* Simulated Top Bar of Desktop App */}
                <div className="flex items-center justify-between pb-3 border-b border-slate-800">
                  <div className="flex items-center gap-2 text-sm">
                    <span className="text-slate-400">Пользователь:</span>
                    <span className="font-semibold text-emerald-400">romosol2012@gmail.com</span>
                    <span className="text-xs px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">Активен</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <div className="px-2.5 py-1 bg-slate-800 rounded-lg text-xs font-medium text-slate-200 border border-slate-700 flex items-center gap-1.5">
                      <FolderOpen className="w-3.5 h-3.5 text-emerald-400" />
                      📂 Открыть list.xlsx
                    </div>
                    <div className="px-2.5 py-1 bg-slate-800 rounded-lg text-xs font-medium text-slate-200 border border-slate-700 flex items-center gap-1.5">
                      <Layers className="w-3.5 h-3.5 text-blue-400" />
                      📋 Открыть results_log.txt
                    </div>
                    <div className="px-2.5 py-1 bg-slate-800 rounded-lg text-xs font-medium text-slate-200 border border-slate-700 flex items-center gap-1.5">
                      <Settings className="w-3.5 h-3.5 text-slate-400" />
                      ⚙ Открыть config.json
                    </div>
                  </div>
                </div>

                {/* Real-time Server Monitor Widget */}
                <div className="bg-slate-950/80 border border-blue-500/30 rounded-xl p-3.5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-inner">
                  <div className="space-y-1">
                    <div className="flex items-center gap-2">
                      <span className="relative flex h-2.5 w-2.5">
                        <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75"></span>
                        <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500"></span>
                      </span>
                      <span className="text-xs font-bold text-slate-200">
                        Отклик сервера: <span className={mockPing.color}>{mockPing.ms} мс</span>
                      </span>
                      <span className="text-[11px] px-2 py-0.5 rounded bg-blue-500/10 border border-blue-500/20 text-blue-300 font-medium">
                        Нагрузка: {mockPing.label}
                      </span>
                    </div>
                    <p className="text-[11px] text-slate-400">
                      Скорость отправки: <span className="text-slate-200 font-mono">0.83 запр/сек</span> (пауза {mockPing.delay}с) | Режим: <span className="text-emerald-400 font-medium">Защита от перегрузки (строго ≤ 1 запр/сек)</span>
                    </p>
                  </div>
                  <button
                    onClick={handleSimulatePing}
                    disabled={mockPing.isChecking}
                    className="px-3 py-1.5 bg-blue-600/20 hover:bg-blue-600/30 border border-blue-500/40 text-blue-300 text-xs font-medium rounded-lg flex items-center gap-1.5 transition-colors whitespace-nowrap"
                  >
                    <Activity className={`w-3.5 h-3.5 ${mockPing.isChecking ? 'animate-spin' : ''}`} />
                    {mockPing.isChecking ? 'Замер...' : '🔄 Проверить отклик'}
                  </button>
                </div>

                {/* Simulated Mode Tabs */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {/* Mode 1: NavAdd */}
                  <div className="bg-slate-950/60 border border-slate-800 rounded-xl p-4 space-y-3.5">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-blue-400 uppercase tracking-wider">Режим 1: NavAdd</span>
                      <span className="text-[11px] text-slate-500">Пакетная запись</span>
                    </div>
                    <div className="space-y-2 text-xs">
                      <div>
                        <label className="block text-slate-400 mb-1">Название мероприятия:</label>
                        <div className="w-full bg-slate-900 border border-slate-700 rounded px-3 py-2 text-slate-200 font-mono text-[11px]">
                          Мастер-класс по анимации в ДОЛ Горный воздух
                        </div>
                      </div>
                      <div>
                        <label className="block text-slate-400 mb-1">Дата и время участия:</label>
                        <div className="w-full bg-slate-900 border border-slate-700 rounded px-3 py-2 text-slate-200 font-mono text-[11px]">
                          2026-08-31 11:00:00
                        </div>
                      </div>
                      <div className="pt-2">
                        <div className="w-full bg-blue-600 rounded-lg py-2 text-center font-semibold text-white text-xs shadow-sm flex items-center justify-center gap-1.5">
                          <Play className="w-3.5 h-3.5 fill-current" />
                          ▶ Запустить запись детей из таблицы
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Mode 2: NavConfirm */}
                  <div className="bg-slate-950/60 border border-slate-800 rounded-xl p-4 space-y-3.5">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-emerald-400 uppercase tracking-wider">Режим 2: Подтверждение и Участие</span>
                      <span className="text-[11px] text-slate-500">approve & participant</span>
                    </div>
                    <div className="space-y-2 text-xs">
                      <div>
                        <label className="block text-slate-400 mb-1">Название мероприятия:</label>
                        <div className="w-full bg-slate-900 border border-slate-700 rounded px-3 py-2 text-slate-200 font-mono text-[11px]">
                          Мастер-класс по анимации в ДОЛ Горный воздух
                        </div>
                      </div>
                      <div className="pt-2 space-y-2">
                        <div className="w-full bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg py-2 text-center font-medium text-slate-200 text-xs">
                          🔍 Проверить статус заявок мероприятия
                        </div>
                        <div className="grid grid-cols-2 gap-2">
                          <div className="w-full bg-emerald-600/30 border border-emerald-500/40 rounded-lg py-2 px-1 text-center font-semibold text-emerald-300 text-[11px] leading-tight">
                            ✓ 1. Подтвердить (approve)
                          </div>
                          <div className="w-full bg-cyan-600/30 border border-cyan-500/40 rounded-lg py-2 px-1 text-center font-semibold text-cyan-300 text-[11px] leading-tight">
                            🎖️ 2. Отметить участие (participant)
                          </div>
                        </div>
                        <p className="text-[10px] text-slate-500 italic text-center pt-0.5">
                          Сначала подтверждение (1), затем отметка участия подтвержденных (2)
                        </p>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Simulated Log Output */}
                <div className="bg-slate-950 rounded-xl p-3 border border-slate-800 font-mono text-[11px] space-y-1 text-slate-300">
                  <div className="text-slate-500 border-b border-slate-800/80 pb-1 flex justify-between">
                    <span>Журнал операций (Консоль)</span>
                    <span className="text-emerald-400">● В сети</span>
                  </div>
                  <div className="text-cyan-400">[ПОИСК] Мероприятие найдено: ID 44220</div>
                  <div className="text-slate-400">  • Неподтвержденных (initial): 0 | Подтвержденных (approve): 1 | Участников (participant): 14</div>
                  <div className="text-cyan-400">[СТАРТ] Массовая отметка участия (state -&gt; participant)...</div>
                  <div className="text-emerald-400 font-medium">  ✓ Заявка #330436: участие успешно отмечено (state -&gt; participant)</div>
                  <div className="text-cyan-300">ИТОГИ: Успешно отмечено участие для 1 заявки!</div>
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === 'guide' && (
          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 space-y-6">
            <div>
              <h2 className="text-base font-bold text-white flex items-center gap-2">
                <Terminal className="w-5 h-5 text-blue-400" />
                Как запустить программу на вашем компьютере
              </h2>
              <p className="text-xs text-slate-400 mt-1">
                Все три файла (приложение, таблица и конфиг) должны лежать в одной папке на компьютере.
              </p>
            </div>

            <div className="space-y-4">
              <div className="bg-slate-950 border border-slate-800 rounded-xl p-4">
                <div className="text-xs font-semibold text-slate-200 mb-2">Шаг 1. Скачайте файлы в одну папку:</div>
                <div className="flex flex-wrap gap-2 text-xs">
                  <button onClick={handleDownloadApp} className="px-3 py-1.5 bg-blue-600/20 text-blue-300 border border-blue-500/30 rounded flex items-center gap-1.5">
                    <Download className="w-3.5 h-3.5" /> 1. navigator_app.py
                  </button>
                  <button onClick={handleDownloadExcel} className="px-3 py-1.5 bg-emerald-600/20 text-emerald-300 border border-emerald-500/30 rounded flex items-center gap-1.5">
                    <Download className="w-3.5 h-3.5" /> 2. list.xlsx (3 столбца)
                  </button>
                  <button onClick={handleDownloadConfig} className="px-3 py-1.5 bg-slate-800 text-slate-300 border border-slate-700 rounded flex items-center gap-1.5">
                    <Download className="w-3.5 h-3.5" /> 3. config.json
                  </button>
                </div>
              </div>

              <div className="bg-slate-950 border border-slate-800 rounded-xl p-4">
                <div className="flex items-center justify-between mb-2">
                  <div className="text-xs font-semibold text-slate-200">Шаг 2. Запустите через командную строку (Терминал / CMD):</div>
                  <button
                    onClick={handleCopyCommand}
                    className="text-xs text-blue-400 hover:text-blue-300 flex items-center gap-1"
                  >
                    {copiedCmd ? <Check className="w-3.5 h-3.5 text-emerald-400" /> : <Copy className="w-3.5 h-3.5" />}
                    {copiedCmd ? 'Скопировано!' : 'Копировать команду'}
                  </button>
                </div>
                <pre className="bg-slate-900 p-3 rounded-lg text-xs font-mono text-emerald-400 border border-slate-800">
pip install requests openpyxl{"\n"}python navigator_app.py
                </pre>
              </div>

              <div className="bg-slate-950 border border-slate-800 rounded-xl p-4 space-y-2 text-xs text-slate-300">
                <div className="font-semibold text-slate-200">Шаг 3. Как пользоваться (Запись, Подтверждение и Участие):</div>
                <ul className="list-disc list-inside space-y-1.5 text-slate-400">
                  <li><b>Режим 1 (Пакетная запись):</b> Введите название мероприятия и дату. Программа проверит <code className="text-emerald-400">is_approved: true</code>, сверит дату рождения и запишет. После успеха в 3-й столбец таблицы автоматически запишется отметка, чтобы исключить дублирование при повторном запуске.</li>
                  <li><b>Режим 2 (Подтверждение и Отметка участия):</b>
                    <ul className="list-circle list-inside pl-4 pt-1 space-y-1 text-slate-400">
                      <li>Нажмите <b>«🔍 Проверить статус заявок»</b>: программа покажет количество неподтвержденных (<code className="text-amber-400">initial</code>), подтвержденных (<code className="text-cyan-400">approve</code>) и участников (<code className="text-emerald-400">participant</code>).</li>
                      <li>Нажмите <b>«✓ 1. Подтвердить заявки»</b>: программа подтвердит все новые заявки (<code className="text-emerald-400">state: approve</code>).</li>
                      <li>Нажмите отдельную кнопку <b>«🎖️ 2. Отметить участие»</b>: программа сделает вызов <code className="text-cyan-400">/api/setActivityOrderState</code> с <code className="text-cyan-400">state: participant</code> строго для тех заявок, которые уже были подтверждены.</li>
                    </ul>
                  </li>
                </ul>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
