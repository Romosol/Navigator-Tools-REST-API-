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
  Gauge,
  GraduationCap,
  Search,
  AlertCircle,
  BookOpen,
  RefreshCw,
  Bug
} from 'lucide-react';

export default function App() {
  const [copiedCode, setCopiedCode] = useState(false);
  const [copiedCmd, setCopiedCmd] = useState(false);
  const [activeTab, setActiveTab] = useState<'overview' | 'code' | 'guide'>('overview');
  const [guiTab, setGuiTab] = useState<'events' | 'confirm' | 'program' | 'study'>('study');
  const [showUpdateModal, setShowUpdateModal] = useState(false);
  const [showBugModal, setShowBugModal] = useState(false);
  const [bugTitle, setBugTitle] = useState('Ошибка при проверке заявки');
  const [bugSection, setBugSection] = useState('Вкладка 4: Сверка со study_list, подтверждение и обучение');
  const [bugDesc, setBugDesc] = useState('1. Что делали: Нажали кнопку «Загрузить список заявок программы»\n2. Что ожидали: Заявки сопоставятся с файлом Excel\n3. Что произошло: Ошибка таймаута при медленном интернете');
  const [bugDiag, setBugDiag] = useState(true);
  const [bugCopied, setBugCopied] = useState(false);
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

  const handleDownloadFile = async (fileName: string) => {
    try {
      const res = await fetch(`/${fileName}?t=${Date.now()}`);
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (e) {
      const a = document.createElement('a');
      a.href = `/${fileName}?t=${Date.now()}`;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }
  };

  const handleDownloadApp = () => handleDownloadFile('navigator_app.py');
  const handleDownloadConfigExample = () => handleDownloadFile('config.example.json');
  const handleDownloadLicense = () => handleDownloadFile('LICENSE');
  const handleDownloadEventExcel = () => handleDownloadFile('event_list.xlsx');
  const handleDownloadProgramExcel = () => handleDownloadFile('programm_list.xlsx');
  const handleDownloadStudyExcel = () => handleDownloadFile('study_list.xlsx');
  const handleDownloadConfig = () => handleDownloadFile('config.example.json');
  const handleDownloadLog = () => handleDownloadFile('logs/results_log.txt');

  const handleDownloadBat = () => {
    const a = document.createElement('a');
    a.href = '/build_exe.bat';
    a.download = 'build_exe.bat';
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
              Графический интерфейс, авто-вход, поддержка буфера обмена в RU/EN раскладках, раздельные таблицы event_list.xlsx и programm_list.xlsx
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={handleDownloadEventExcel}
            className="px-3 py-1.5 rounded-lg border border-emerald-500/30 bg-emerald-500/10 hover:bg-emerald-500/20 text-emerald-300 text-xs font-medium flex items-center gap-1.5 transition-colors"
            title="Таблица для мероприятий"
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-400" />
            event_list.xlsx
          </button>
          <button
            onClick={handleDownloadProgramExcel}
            className="px-3 py-1.5 rounded-lg border border-purple-500/30 bg-purple-500/10 hover:bg-purple-500/20 text-purple-300 text-xs font-medium flex items-center gap-1.5 transition-colors"
            title="Таблица для зачисления на программы (3 столбца)"
          >
            <GraduationCap className="w-3.5 h-3.5 text-purple-400" />
            programm_list.xlsx
          </button>
          <button
            onClick={handleDownloadStudyExcel}
            className="px-3 py-1.5 rounded-lg border border-sky-500/30 bg-sky-500/10 hover:bg-sky-500/20 text-sky-300 text-xs font-medium flex items-center gap-1.5 transition-colors"
            title="Таблица для подтверждения и обучения (2 столбца: ФИО и флаг зачисления)"
          >
            <FileSpreadsheet className="w-3.5 h-3.5 text-sky-400" />
            study_list.xlsx
          </button>
          <button
            onClick={handleDownloadLog}
            className="px-2.5 py-1.5 rounded-lg border border-blue-500/30 bg-blue-500/10 hover:bg-blue-500/20 text-blue-300 text-xs font-medium flex items-center gap-1.5 transition-colors"
            title="Логи хранятся исключительно в папке logs/"
          >
            <Layers className="w-3.5 h-3.5 text-blue-400" />
            logs/results_log.txt
          </button>
          <button
            onClick={handleDownloadConfig}
            className="px-2.5 py-1.5 rounded-lg border border-slate-700 hover:bg-slate-800 text-slate-300 text-xs font-medium flex items-center gap-1.5 transition-colors"
          >
            <Settings className="w-3.5 h-3.5 text-slate-400" />
            config.json
          </button>
          <a
            href="https://github.com/Romosol/Navigator-Tools-REST-API-/releases/latest"
            target="_blank"
            rel="noreferrer"
            className="px-3 py-1.5 rounded-lg border border-blue-500/40 bg-blue-600/20 hover:bg-blue-600/30 text-blue-300 text-xs font-semibold flex items-center gap-1.5 transition-colors"
            title="Перейти к скачиванию готового NavigatorApp.exe"
          >
            <Download className="w-3.5 h-3.5 text-blue-400" />
            Релиз NavigatorApp.exe
          </a>
          <button
            onClick={handleDownloadApp}
            className="px-3.5 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold flex items-center gap-1.5 transition-all shadow-md shadow-blue-600/20"
            title="Скачать последнюю версию navigator_app.py со всеми обновлениями (кнопка СТОП, авто-даты)"
          >
            <Download className="w-3.5 h-3.5" />
            navigator_app.py (актуальная v2.1)
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
            {/* 5 Feature Highlights */}
            <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-3.5">
              <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4 flex flex-col justify-between space-y-3">
                <div className="flex items-start gap-3">
                  <div className="p-2 rounded-lg bg-purple-500/10 text-purple-400 border border-purple-500/20">
                    <GraduationCap className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-sm font-semibold text-slate-200">Зачисление на программу</h2>
                    <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                      Новая вкладка: создание заявок через <code>/api/rest/order</code> по <b>event_id</b>, <b>group_id</b> и <b>academic_year_id</b>.
                    </p>
                  </div>
                </div>
                <div className="text-[11px] text-purple-400/90 bg-purple-950/30 p-2 rounded border border-purple-500/20 flex items-center gap-1">
                  <CheckCircle2 className="w-3.5 h-3.5" />
                  Опции сертификатов
                </div>
              </div>

              <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4 flex flex-col justify-between space-y-3">
                <div className="flex items-start gap-3">
                  <div className="p-2 rounded-lg bg-blue-500/10 text-blue-400 border border-blue-500/20">
                    <Calendar className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-sm font-semibold text-slate-200">Запись на мероприятие</h2>
                    <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                      Поиск мероприятия, сверка тезок по дате рождения и запись через <code>/api/rest/activityOrder</code>.
                    </p>
                  </div>
                </div>
                <div className="text-[11px] text-blue-400/90 bg-blue-950/30 p-2 rounded border border-blue-500/20">
                  Сохранение в config.json
                </div>
              </div>

              <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4 flex flex-col justify-between space-y-3">
                <div className="flex items-start gap-3">
                  <div className="p-2 rounded-lg bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">
                    <CheckCircle2 className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-sm font-semibold text-slate-200">Подтверждение и Участие</h2>
                    <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                      Двухэтапная обработка: кнопка 1 переводит в <code>approve</code>, кнопка 2 отмечает <code>participant</code>.
                    </p>
                  </div>
                </div>
                <div className="text-[11px] text-emerald-400/90 bg-emerald-950/30 p-2 rounded border border-emerald-500/20">
                  Раздельные вызовы API
                </div>
              </div>

              <div className="bg-slate-900/80 border border-slate-800 rounded-xl p-4 flex flex-col justify-between space-y-3">
                <div className="flex items-start gap-3">
                  <div className="p-2 rounded-lg bg-amber-500/10 text-amber-400 border border-amber-500/20">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div>
                    <h2 className="text-sm font-semibold text-slate-200">Таблицы Excel & Сверка ФИО</h2>
                    <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                      <code>study_list.xlsx</code> (2 столбца: ФИО и флаг) сверяет Фамилию и Имя, защищая от чужих заявок. Инициалы отмечаются как ошибка!
                    </p>
                  </div>
                </div>
                <div className="text-[11px] text-amber-400/90 bg-amber-950/30 p-2 rounded border border-amber-500/20 flex items-center justify-between">
                  <span>Сверка Фамилия + Имя</span>
                  <span className="text-[10px] text-emerald-400">study_list.xlsx</span>
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
                      Строго <b>≤ 1 запрос в сек</b> (пауза 1.2 с). Адаптивное замедление и мониторинг отклика сервера.
                    </p>
                  </div>
                </div>
                <div className="text-[11px] text-indigo-400/90 bg-indigo-950/30 p-2 rounded border border-indigo-500/20 flex items-center gap-1.5">
                  <Activity className="w-3.5 h-3.5" />
                  Пинг в реальном времени
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
                  <span className="ml-2 font-medium text-slate-300">Навигатор: Автоматизация (Мероприятия, Подтверждение & Зачисление) — Окно программы</span>
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
                  <div className="flex flex-wrap items-center gap-2">
                    <div className="px-2.5 py-1 bg-slate-800 rounded-lg text-xs font-medium text-slate-200 border border-slate-700 flex items-center gap-1.5" title="Таблица мероприятий">
                      <FolderOpen className="w-3.5 h-3.5 text-emerald-400" />
                      📊 event_list.xlsx
                    </div>
                    <div className="px-2.5 py-1 bg-slate-800 rounded-lg text-xs font-medium text-slate-200 border border-slate-700 flex items-center gap-1.5" title="Таблица программ">
                      <FolderOpen className="w-3.5 h-3.5 text-purple-400" />
                      🎓 programm_list.xlsx
                    </div>
                    <div className="px-2.5 py-1 bg-slate-800 rounded-lg text-xs font-medium text-slate-200 border border-slate-700 flex items-center gap-1.5" title="Таблица обучения">
                      <FolderOpen className="w-3.5 h-3.5 text-sky-400" />
                      📑 study_list.xlsx
                    </div>
                    <div className="px-2.5 py-1 bg-slate-800 rounded-lg text-xs font-medium text-slate-200 border border-slate-700 flex items-center gap-1.5" title="Уникальный лог текущего запуска с датой и временем">
                      <Layers className="w-3.5 h-3.5 text-blue-400" />
                      📋 Лог (results_log_2026-09-14_11-15-30.txt)
                    </div>
                    <div className="px-2.5 py-1 bg-slate-800 rounded-lg text-xs font-medium text-slate-200 border border-slate-700 flex items-center gap-1.5" title="Папка со всеми архивными логами">
                      <FolderOpen className="w-3.5 h-3.5 text-amber-400" />
                      📁 Папка логов (logs/)
                    </div>
                    <div className="px-2.5 py-1 bg-slate-800 rounded-lg text-xs font-medium text-slate-200 border border-slate-700 flex items-center gap-1.5">
                      <Settings className="w-3.5 h-3.5 text-slate-400" />
                      ⚙ config.json
                    </div>
                    <button
                      onClick={() => setShowUpdateModal(true)}
                      className="px-2.5 py-1 bg-blue-600/20 hover:bg-blue-600/40 text-blue-300 border border-blue-500/40 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors cursor-pointer"
                      title="Проверить обновления приложения"
                    >
                      <RefreshCw className="w-3.5 h-3.5 text-blue-400" />
                      🔄 Обновления (v2.1.0)
                    </button>
                    <button
                      onClick={() => setShowBugModal(true)}
                      className="px-2.5 py-1 bg-red-600/20 hover:bg-red-600/40 text-red-300 border border-red-500/40 rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors cursor-pointer"
                      title="Сообщить об ошибке в GitHub Issues"
                    >
                      <Bug className="w-3.5 h-3.5 text-red-400" />
                      🐞 Баг-репорт
                    </button>
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

                {/* Notebook Tabs Bar */}
                <div className="flex items-center gap-1 bg-slate-950 p-1 rounded-xl border border-slate-800">
                  <button
                    onClick={() => setGuiTab('events')}
                    className={`flex-1 py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-2 transition-all ${
                      guiTab === 'events'
                        ? 'bg-blue-600 text-white shadow-md'
                        : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
                    }`}
                  >
                    <Calendar className="w-3.5 h-3.5" />
                    📝 1. Запись на мероприятие
                  </button>
                  <button
                    onClick={() => setGuiTab('confirm')}
                    className={`flex-1 py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-2 transition-all ${
                      guiTab === 'confirm'
                        ? 'bg-emerald-600 text-white shadow-md'
                        : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
                    }`}
                  >
                    <CheckCircle2 className="w-3.5 h-3.5" />
                    ✓ 2. Подтверждение и участие
                  </button>
                  <button
                    onClick={() => setGuiTab('program')}
                    className={`flex-1 py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-2 transition-all ${
                      guiTab === 'program'
                        ? 'bg-purple-600 text-white shadow-md'
                        : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
                    }`}
                  >
                    <GraduationCap className="w-3.5 h-3.5" />
                    🎓 3. Зачисление на программу
                  </button>
                  <button
                    onClick={() => setGuiTab('study')}
                    className={`flex-1 py-2 px-3 rounded-lg text-xs font-semibold flex items-center justify-center gap-2 transition-all ${
                      guiTab === 'study'
                        ? 'bg-sky-600 text-white shadow-md'
                        : 'text-slate-400 hover:text-slate-200 hover:bg-slate-900'
                    }`}
                  >
                    <BookOpen className="w-3.5 h-3.5" />
                    📖 4. Подтверждение и обучение
                  </button>
                </div>

                {/* Simulated Mode Body */}
                {guiTab === 'study' && (
                  <div className="bg-slate-950/60 border border-slate-800 rounded-xl p-4 space-y-3.5">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-sky-400 uppercase tracking-wider">Вкладка 4: Подтверждение заявок и отметка об обучении</span>
                      <span className="text-[11px] text-sky-300/80 font-mono">POST /api/approveRequest &amp; POST /api/studyRequest</span>
                    </div>

                    {/* Program Search Field */}
                    <div className="space-y-1">
                      <label className="block text-slate-400 text-xs">Название программы (GET /api/rest/events):</label>
                      <div className="flex gap-2">
                        <div className="flex-1 bg-slate-900 border border-sky-500/40 rounded px-3 py-1.5 text-slate-200 text-xs font-medium truncate">
                          "Мобильный Технопарк VR/AR/IT" (ПДО Габдрахманов Л.И.)
                        </div>
                        <div className="px-3 py-1.5 bg-sky-600/30 border border-sky-500/50 rounded text-sky-300 text-xs font-semibold flex items-center gap-1 cursor-default">
                          <Search className="w-3 h-3" />
                          Проверить
                        </div>
                      </div>
                      <div className="flex items-center gap-2 text-[11px] text-emerald-400 bg-emerald-950/40 border border-emerald-500/20 px-2 py-1 rounded">
                        <CheckCircle2 className="w-3.5 h-3.5" />
                        <span>✓ Программа найдена: <b>ID 12345</b> (Дворец творчества)</span>
                      </div>
                    </div>

                    {/* Group Search Field */}
                    <div className="space-y-1">
                      <label className="block text-slate-400 text-xs">Группа (название или ID):</label>
                      <div className="flex gap-2">
                        <div className="flex-1 bg-slate-900 border border-slate-700 rounded px-3 py-1.5 text-slate-300 text-xs font-medium placeholder-slate-500">
                          Группа 1
                        </div>
                        <div className="px-3 py-1.5 bg-slate-800 border border-slate-700 rounded text-slate-300 text-xs font-semibold flex items-center gap-1 cursor-default">
                          <Search className="w-3 h-3" />
                          Проверить
                        </div>
                      </div>
                      <p className="text-[11px] text-sky-400/90 italic">
                        💡 Можно оставить пустым для зачисления всех групп выбранной программы!
                      </p>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
                      <div>
                        <label className="block text-slate-400 mb-1">Учебный год:</label>
                        <div className="w-full bg-slate-900 border border-slate-700 rounded px-3 py-1.5 text-slate-200 font-mono text-[11px] flex justify-between items-center">
                          <span>2026/2027</span>
                          <span className="text-[10px] text-sky-400 font-sans">→ academic_year_id: 2026</span>
                        </div>
                      </div>
                      <div>
                        <label className="block text-slate-400 mb-1">Приказ о зачислении на обучение:</label>
                        <div className="w-full bg-slate-900 border border-slate-700 rounded px-3 py-1.5 text-slate-200 font-mono text-[11px] flex flex-wrap gap-2 items-center">
                          <span>№ 183</span>
                          <span className="text-slate-500">|</span>
                          <span>от 31.08.2026</span>
                          <span className="text-slate-500">|</span>
                          <span>старт: 01.09.2026</span>
                        </div>
                        <span className="text-[10px] text-emerald-400 font-sans">
                          Форматы: <b>дд.мм.гг</b>, <b>дд.мм.гггг</b> или <b>гггг-мм-дд</b> (Ист. фин. убран)
                        </span>
                      </div>
                    </div>

                    {/* Excel Verification Box */}
                    <div className="bg-slate-900/90 border border-sky-500/30 rounded-lg p-3 space-y-2">
                      <div className="flex items-center justify-between">
                        <label className="flex items-center gap-2 text-xs text-slate-200 cursor-default font-medium">
                          <input type="checkbox" checked readOnly className="rounded text-sky-500 bg-slate-800 border-slate-700 w-3.5 h-3.5 accent-sky-500" />
                          <span>Сверять Фамилию и Имя со study_list.xlsx (отчество игнорируется, инициалы отмечаются как ошибка)</span>
                        </label>
                        <div className="flex gap-1.5">
                          <button onClick={handleDownloadStudyExcel} className="px-2 py-1 bg-sky-950/60 border border-sky-500/30 rounded text-sky-300 text-[10px] flex items-center gap-1 hover:bg-sky-900/40">
                            <FileSpreadsheet className="w-3 h-3" />
                            study_list.xlsx
                          </button>
                        </div>
                      </div>
                      <div className="flex items-center justify-between text-[11px] text-sky-300/90 bg-sky-950/30 px-2.5 py-1.5 rounded border border-sky-500/20">
                        <span>Таблица study_list.xlsx: 15 детей. 2 столбца: ФИО и флаг. Сравнение по Фамилии и Имени активно.</span>
                        <span className="text-[10px] text-emerald-400 font-mono">Совпало: 15 из 15</span>
                      </div>
                      <div className="text-[10px] text-amber-300/90 bg-amber-950/30 px-2 py-1 rounded border border-amber-500/20 flex items-center gap-1.5">
                        <AlertCircle className="w-3.5 h-3.5 text-amber-400 flex-shrink-0" />
                        <span>Если в Excel указаны инициалы ('Иванов И.И.'), система автоматически отметит во 2-м столбце: <i>«⚠️ Неправильный формат (инициалы)»</i> и уведомит пользователя!</span>
                      </div>
                    </div>

                    <div className="pt-1">
                      <div className="w-full bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-lg py-2 text-center font-medium text-slate-200 text-xs flex items-center justify-center gap-1.5 cursor-default">
                        <Search className="w-3.5 h-3.5" />
                        🔍 Найти и проверить заявки программы (GET /api/rest/order)
                      </div>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 pt-1">
                      <div className="bg-emerald-600/20 border border-emerald-500/40 rounded-lg py-2 px-2 text-center text-emerald-300 text-[11px] font-semibold">
                        ✓ 1. Подтвердить 1 заявку из Excel (initial)
                      </div>
                      <div className="bg-sky-600/20 border border-sky-500/40 rounded-lg py-2 px-2 text-center text-sky-300 text-[11px] font-semibold">
                        🎓 2. Зачислить на обучение 1 заявку из Excel (approve)
                      </div>
                      <div className="bg-indigo-600/20 border border-indigo-500/40 rounded-lg py-2 px-2 text-center text-indigo-300 text-[11px] font-semibold">
                        ⚡ 3. Зачислить всё (2 заявки из Excel)
                      </div>
                    </div>
                  </div>
                )}

                {/* Simulated Mode Body */}
                {guiTab === 'events' && (
                  <div className="bg-slate-950/60 border border-slate-800 rounded-xl p-4 space-y-3.5">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-blue-400 uppercase tracking-wider">Вкладка 1: Пакетная запись на мероприятие</span>
                      <span className="text-[11px] text-slate-500">POST /api/rest/activityOrder</span>
                    </div>
                    <div className="space-y-3 text-xs">
                      <div>
                        <label className="block text-slate-400 mb-1">Название мероприятия:</label>
                        <div className="flex gap-2">
                          <div className="flex-1 bg-slate-900 border border-slate-700 rounded px-3 py-2 text-slate-200 font-mono text-[11px]">
                            Мастер-класс по робототехнике
                          </div>
                          <div className="px-3 py-1.5 bg-blue-600/30 border border-blue-500/50 rounded text-blue-300 text-xs font-semibold flex items-center gap-1 cursor-default whitespace-nowrap">
                            <Search className="w-3 h-3" />
                            Проверить
                          </div>
                        </div>
                        <span className="text-[10px] text-slate-500">Сохраняется в config.json</span>
                      </div>
                      <div>
                        <label className="block text-slate-400 mb-1">Дата и время участия:</label>
                        <div className="w-full bg-slate-900 border border-slate-700 rounded px-3 py-2 text-slate-200 font-mono text-[11px]">
                          2026-08-31 11:00:00
                        </div>
                        <span className="text-[10px] text-slate-500">Формат: ГГГГ-ММ-ДД ЧЧ:ММ:СС</span>
                      </div>
                      <div className="pt-1">
                        <div className="w-full bg-blue-600 rounded-lg py-2 text-center font-semibold text-white text-xs shadow-sm flex items-center justify-center gap-1.5 cursor-default">
                          <Play className="w-3.5 h-3.5 fill-current" />
                          ▶ Запустить запись детей из таблицы
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {guiTab === 'confirm' && (
                  <div className="bg-slate-950/60 border border-slate-800 rounded-xl p-4 space-y-3.5">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-emerald-400 uppercase tracking-wider">Вкладка 2: Подтверждение и Участие</span>
                      <span className="text-[11px] text-slate-500">POST /api/setActivityOrderState</span>
                    </div>
                    <div className="space-y-3 text-xs">
                      <div>
                        <label className="block text-slate-400 mb-1">Название мероприятия:</label>
                        <div className="flex gap-2">
                          <div className="flex-1 bg-slate-900 border border-slate-700 rounded px-3 py-2 text-slate-200 font-mono text-[11px]">
                            Мастер-класс по робототехнике
                          </div>
                          <div className="px-3 py-1.5 bg-emerald-600/30 border border-emerald-500/50 rounded text-emerald-300 text-xs font-semibold flex items-center gap-1 cursor-default whitespace-nowrap">
                            <Search className="w-3 h-3" />
                            Проверить
                          </div>
                        </div>
                      </div>
                      <div className="space-y-2 pt-1">
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
                )}

                {guiTab === 'program' && (
                  <div className="bg-slate-950/60 border border-slate-800 rounded-xl p-4 space-y-3.5">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-purple-400 uppercase tracking-wider">Вкладка 3: Зачисление на учебную программу</span>
                      <span className="text-[11px] text-purple-300/80 font-mono">GET /api/rest/events + POST /api/rest/order</span>
                    </div>

                    {/* Program Search Field */}
                    <div className="space-y-1">
                      <label className="block text-slate-400 text-xs">Название программы (поиск через GET /api/rest/events):</label>
                      <div className="flex gap-2">
                        <div className="flex-1 bg-slate-900 border border-purple-500/40 rounded px-3 py-1.5 text-slate-200 text-xs font-medium truncate">
                          "Мобильный Технопарк VR/AR/IT" (ПДО Габдрахманов Л.И.)
                        </div>
                        <div className="px-3 py-1.5 bg-purple-600/30 border border-purple-500/50 rounded text-purple-300 text-xs font-semibold flex items-center gap-1 cursor-default">
                          <Search className="w-3 h-3" />
                          Проверить
                        </div>
                      </div>
                      <div className="flex items-center gap-2 text-[11px] text-emerald-400 bg-emerald-950/40 border border-emerald-500/20 px-2 py-1 rounded">
                        <CheckCircle2 className="w-3.5 h-3.5" />
                        <span>✓ Найдено: <b>ID 12345</b> (Мест в группах: 120 | Организация: Дворец творчества)</span>
                      </div>
                    </div>

                    {/* Group Search Field */}
                    <div className="space-y-1">
                      <label className="block text-slate-400 text-xs">Группа (название или ID, проверка принадлежности программе через GET /api/rest/eventGroups):</label>
                      <div className="flex gap-2">
                        <div className="flex-1 bg-slate-900 border border-purple-500/40 rounded px-3 py-1.5 text-slate-200 text-xs font-medium truncate">
                          Группа 1
                        </div>
                        <div className="px-3 py-1.5 bg-purple-600/30 border border-purple-500/50 rounded text-purple-300 text-xs font-semibold flex items-center gap-1 cursor-default">
                          <Search className="w-3 h-3" />
                          Проверить
                        </div>
                      </div>
                      <div className="flex items-center gap-2 text-[11px] text-emerald-400 bg-emerald-950/40 border border-emerald-500/20 px-2 py-1 rounded">
                        <CheckCircle2 className="w-3.5 h-3.5 flex-shrink-0" />
                        <span>✓ Группа подтверждена для программы <b>ID 12345</b>: «Группа 1» (ID: <b>67890</b> | педагог: Иванов И.И. | мест: 30)</span>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
                      <div>
                        <label className="block text-slate-400 mb-1">ID группы (group_id):</label>
                        <div className="w-full bg-slate-900 border border-slate-700 rounded px-3 py-1.5 text-slate-200 font-mono text-[11px]">
                          67890
                        </div>
                        <span className="text-[10px] text-slate-500">Автоматически сопоставлен с программой</span>
                      </div>
                      <div>
                        <label className="block text-slate-400 mb-1">Учебный год (например, 2026/2027 или 2026):</label>
                        <div className="w-full bg-slate-900 border border-slate-700 rounded px-3 py-1.5 text-slate-200 font-mono text-[11px] flex justify-between items-center">
                          <span>2026/2027</span>
                          <span className="text-[10px] text-purple-400 font-sans">→ academic_year_id: 2026</span>
                        </div>
                        <span className="text-[10px] text-slate-500">Поддерживаются оба формата ввода</span>
                      </div>
                    </div>

                    <div className="flex items-center gap-6 pt-1 text-xs text-slate-300">
                      <label className="flex items-center gap-2 cursor-pointer">
                        <input type="checkbox" defaultChecked className="rounded border-slate-700 text-purple-600 focus:ring-purple-500" readOnly />
                        <span>Создавать сертификат (create_certificate: true)</span>
                      </label>
                      <label className="flex items-center gap-2 cursor-pointer">
                        <input type="checkbox" className="rounded border-slate-700 text-purple-600 focus:ring-purple-500" readOnly />
                        <span>Использовать сертификат (use_certificate)</span>
                      </label>
                    </div>

                    <div className="pt-1">
                      <div className="w-full bg-purple-600 hover:bg-purple-500 rounded-lg py-2.5 text-center font-semibold text-white text-xs shadow-md shadow-purple-600/20 flex items-center justify-center gap-1.5 cursor-default">
                        <Play className="w-3.5 h-3.5 fill-current" />
                        ▶ Запустить зачисление детей на программу из таблицы (programm_list.xlsx)
                      </div>
                    </div>
                  </div>
                )}

                {/* Simulated Big Red Abort Button */}
                <div className="pt-1">
                  <div className="w-full bg-red-600 hover:bg-red-700 active:bg-red-800 text-white font-bold py-2.5 px-4 rounded-xl text-center shadow-lg shadow-red-900/30 border-2 border-red-500 flex items-center justify-center gap-2 cursor-pointer select-none transition-all">
                    <span className="text-base">🛑</span>
                    <span className="text-xs tracking-wide uppercase">ПРЕРВАТЬ ТЕКУЩУЮ ЗАДАЧУ И ВЕРНУТЬСЯ В ИСХОДНОЕ СОСТОЯНИЕ</span>
                  </div>
                </div>

                {/* Simulated Log Output */}
                <div className="bg-slate-950 rounded-xl p-3 border border-slate-800 font-mono text-[11px] space-y-1 text-slate-300">
                  <div className="text-slate-500 border-b border-slate-800/80 pb-1 flex justify-between">
                    <span>Журнал операций (Консоль)</span>
                    <span className="text-emerald-400">● В сети</span>
                  </div>
                  {guiTab === 'study' ? (
                    <>
                      <div className="text-cyan-400">[ТАБЛИЦА] Таблица study_list.xlsx: 15 детей (подтверждено: 1, зачислено: 13). 2 столбца: ФИО и флаг.</div>
                      <div className="text-cyan-400">[ЗАПРОС ЗАЯВОК] Программа: 'Основы программирования и робототехники', группа: 'Группа 1', год: 2026...</div>
                      <div className="text-emerald-400 font-medium">✓ Заявки программы 'Основы программирования и робототехники' (ID: 12345) успешно загружены:</div>
                      <div className="text-slate-300">   • Всего найдено заявок на сервере: 15 | Группа: «Группа 1» (ID: 67890)</div>
                      <div className="text-cyan-300">   • 📑 Сверка со study_list.xlsx (15 детей в файле):</div>
                      <div className="text-emerald-400">      ✓ Сравнение по Фамилии и Имени (отчество в Excel игнорируется): 15 из 15 совпало</div>
                      <div className="text-yellow-400">   • ⏳ Неподтвержденные (initial) — готовы к подтверждению: 1</div>
                      <div className="text-sky-300">   • 📋 Подтвержденные (approve) — готовы к зачислению на обучение: 1</div>
                      <div className="text-emerald-400">   • 🎓 Обучаются (study) — уже зачислены приказом: 13</div>
                      <div className="text-slate-400">  [01] Заявка #3484354 | Смирнов Алексей Дмитриевич     | ДР: 14.07.2011 | ✓ Excel (стр.2: 'Смирнов Алексей Дмитриевич') | Статус: 'initial'</div>
                      <div className="text-slate-400">  [02] Заявка #3484210 | Васильева Екатерина Сергеевна  | ДР: 02.03.2012 | ✓ Excel (стр.3: 'Васильева Екатерина') | Статус: 'approve'</div>
                      <div className="text-cyan-400">[СТАРТ] Шаг 1: Подтверждение заявок initial ➜ approve...</div>
                      <div className="text-emerald-400 font-medium">  ✓ Заявка #3484354 (Смирнов Алексей): подтверждена (approve) ➜ флаг в Excel обновлен</div>
                      <div className="text-cyan-400">[СТАРТ] Шаг 2: Зачисление на обучение approve ➜ study (приказ №1 от 2026-08-31)...</div>
                      <div className="text-emerald-400 font-medium">  ✓ Заявка #3484354: зачислена на обучение (study, приказ №1) ➜ флаг в Excel: 'Зачислен'</div>
                      <div className="text-emerald-400 font-medium">  ✓ Заявка #3484210: зачислена на обучение (study, приказ №1) ➜ флаг в Excel: 'Зачислен'</div>
                      <div className="text-sky-300">ИТОГИ: Успешно зачислено 2 детей на обучение (study, приказ №1 от 2026-08-31)!</div>
                    </>
                  ) : guiTab === 'program' ? (
                    <>
                      <div className="text-cyan-400">[ПОИСК] Поиск программы: 'Основы программирования и робототехники'...</div>
                      <div className="text-emerald-400 font-medium">✓ Программа найдена: 'Основы программирования и робототехники' (ID: 12345)</div>
                      <div className="text-cyan-400">[ПРОВЕРКА] Проверка принадлежности группы 'Группа 1' к программе ID 12345...</div>
                      <div className="text-emerald-400 font-medium">✓ Группа подтверждена для программы: 'Группа 1' (ID: 67890, педагог: Иванов И.И., мест: 30)</div>
                      <div className="text-cyan-400">[СТАРТ] Зачисление на программу (event_id: 12345, group_id: 67890, год: 2026, файл: programm_list.xlsx)...</div>
                      <div className="text-slate-400">Загружено записей из таблицы (programm_list.xlsx): 15</div>
                      <div className="text-slate-300">[1] Иванов Иван Иванович (ДР: 15.05.2012)</div>
                      <div className="text-cyan-300">   [ПРОВЕРКА ДУБЛИКАТА] Поиск существующих заявок (state_grid in initial, approve, study)...</div>
                      <div className="text-emerald-400 font-medium">   ✓ УСПЕШНО: Заявка на зачисление создана (#3485996) (Kid ID: f6c7c6d2..., Parent: 791100)</div>
                      <div className="text-cyan-300">   💾 Статус в programm_list.xlsx сохранен: 'Зачислен (Заявка #3485996)' (строка 2)</div>
                      <div className="text-slate-300">[2] Петров Петр Петрович (ДР: 10.05.2010)</div>
                      <div className="text-cyan-400">   ⏭️ ПРОПУСК: Заявка уже зарегистрирована в Навигаторе (#3481200, статус: 'study')</div>
                      <div className="text-purple-300">ИТОГИ: Успешно создано новых заявок: 14 | Пропущено дубликатов: 1 | Отчет сохранен в: logs/results_log_2026-09-14_11-15-30.txt</div>
                    </>
                  ) : guiTab === 'confirm' ? (
                    <>
                      <div className="text-cyan-400">[ПОИСК] Мероприятие найдено: ID 44220</div>
                      <div className="text-slate-400">  • Неподтвержденных (initial): 0 | Подтвержденных (approve): 1 | Участников (participant): 14</div>
                      <div className="text-cyan-400">[СТАРТ] Массовая отметка участия (state -&gt; participant)...</div>
                      <div className="text-emerald-400 font-medium">  ✓ Заявка #330436: участие успешно отмечено (state -&gt; participant)</div>
                      <div className="text-cyan-300">ИТОГИ: Успешно отмечено участие для 1 заявки!</div>
                    </>
                  ) : (
                    <>
                      <div className="text-cyan-400">[СТАРТ] Пакетная запись детей на мероприятие (event_list.xlsx)...</div>
                      <div className="text-slate-400">Загружено записей из таблицы (event_list.xlsx): 15</div>
                      <div className="text-slate-300">[1] Иванов Иван Иванович (ДР: 15.03.2012)</div>
                      <div className="text-emerald-400 font-medium">   ✓ УСПЕШНО: Заявка создана (#330440)</div>
                      <div className="text-cyan-300">   💾 Статус в event_list.xlsx сохранен: 'Добавлен (Заявка #330440)' (строка 2)</div>
                      <div className="text-blue-300">ИТОГИ: Успешно добавлено: 15 детей.</div>
                    </>
                  )}
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
                Все файлы (приложение, таблицы и конфиг) должны лежать в одной папке на компьютере.
              </p>
            </div>

            <div className="space-y-4">
              <div className="bg-slate-950 border border-slate-800 rounded-xl p-4">
                <div className="text-xs font-semibold text-slate-200 mb-2">Шаг 1. Скачайте файлы в одну папку:</div>
                <div className="flex flex-wrap gap-2 text-xs">
                  <a
                    href="https://github.com/Romosol/Navigator-Tools-REST-API-/releases/latest"
                    target="_blank"
                    rel="noreferrer"
                    className="px-3 py-1.5 bg-blue-600 text-white rounded flex items-center gap-1.5 hover:bg-blue-500 transition-colors font-semibold shadow-sm"
                  >
                    <Download className="w-3.5 h-3.5" /> 🚀 Скачать NavigatorApp.exe (Релиз)
                  </a>
                  <button onClick={handleDownloadApp} className="px-3 py-1.5 bg-blue-600/20 text-blue-300 border border-blue-500/30 rounded flex items-center gap-1.5 hover:bg-blue-600/30 transition-colors font-medium">
                    <Download className="w-3.5 h-3.5" /> navigator_app.py (Исходный код)
                  </button>
                  <button onClick={handleDownloadEventExcel} className="px-3 py-1.5 bg-emerald-600/20 text-emerald-300 border border-emerald-500/30 rounded flex items-center gap-1.5" title="Таблица для мероприятий">
                    <Download className="w-3.5 h-3.5" /> event_list.xlsx
                  </button>
                  <button onClick={handleDownloadProgramExcel} className="px-3 py-1.5 bg-purple-600/20 text-purple-300 border border-purple-500/30 rounded flex items-center gap-1.5" title="Таблица для программ">
                    <Download className="w-3.5 h-3.5" /> programm_list.xlsx
                  </button>
                  <button onClick={handleDownloadStudyExcel} className="px-3 py-1.5 bg-sky-600/20 text-sky-300 border border-sky-500/30 rounded flex items-center gap-1.5" title="Таблица для сверки обучения">
                    <Download className="w-3.5 h-3.5" /> study_list.xlsx
                  </button>
                  <button onClick={handleDownloadConfigExample} className="px-3 py-1.5 bg-slate-800 text-slate-300 border border-slate-700 rounded flex items-center gap-1.5">
                    <Download className="w-3.5 h-3.5" /> config.example.json
                  </button>
                  <button onClick={handleDownloadLicense} className="px-3 py-1.5 bg-slate-800 text-slate-300 border border-slate-700 rounded flex items-center gap-1.5">
                    <Download className="w-3.5 h-3.5" /> LICENSE (MIT)
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

              <div className="bg-slate-950 border border-slate-800 rounded-xl p-4 space-y-3 text-xs text-slate-300">
                <div className="font-semibold text-slate-200">Шаг 3. Как пользоваться 4 режимами программы:</div>
                <ul className="list-disc list-inside space-y-2.5 text-slate-400">
                  <li>
                    <b className="text-blue-300">Вкладка 1 (Запись на мероприятие):</b> Использует таблицу <code className="text-emerald-400">event_list.xlsx</code>. Введите название мероприятия и дату. Программа проверит <code className="text-emerald-400">is_approved: true</code>, сверит дату рождения и создаст заявку через <code className="text-blue-300">/api/rest/activityOrder</code>. После успеха в 3-й столбец таблицы автоматически запишется <code className="text-cyan-300">Добавлен (Заявка #...)</code>.
                  </li>
                  <li>
                    <b className="text-emerald-300">Вкладка 2 (Подтверждение и Отметка участия):</b>
                    <ul className="list-circle list-inside pl-4 pt-1 space-y-1 text-slate-400">
                      <li>Нажмите <b>«🔍 Проверить статус заявок»</b>: программа покажет количество неподтвержденных (<code className="text-amber-400">initial</code>), подтвержденных (<code className="text-cyan-400">approve</code>) и участников (<code className="text-emerald-400">participant</code>).</li>
                      <li>Нажмите <b>«✓ 1. Подтвердить заявки»</b>: программа подтвердит все новые заявки (<code className="text-emerald-400">state: approve</code>).</li>
                      <li>Нажмите отдельную кнопку <b>«🎖️ 2. Отметить участие»</b>: программа сделает вызов <code className="text-cyan-400">/api/setActivityOrderState</code> с <code className="text-cyan-400">state: participant</code> строго для тех заявок, которые уже были подтверждены.</li>
                    </ul>
                  </li>
                  <li>
                    <b className="text-purple-300">Вкладка 3 (Зачисление на учебную программу):</b>
                    <ul className="list-circle list-inside pl-4 pt-1 space-y-1 text-slate-400">
                      <li>Использует отдельную таблицу <code className="text-purple-400">programm_list.xlsx</code> (ФИО, Дата рождения, Статус).</li>
                      <li>Введите <b className="text-slate-300">Название программы</b> (например, <code>Основы программирования и робототехники</code>). Программа автоматически найдет её через <code className="text-purple-300">GET /api/rest/events</code>.</li>
                      <li>Укажите <b className="text-slate-300">Группу</b> по названию (например, <code>Группа 1</code>) или точный ID (<code>67890</code>). Программа проверяет принадлежность группы программе через <code className="text-purple-300">GET /api/rest/eventGroups?extFilters=[{'{'}"property":"event_id","value":...{'}'}]</code>.</li>
                      <li>Укажите <b className="text-slate-300">Учебный год</b>: поддерживаются оба формата — как на сайте <code>2026/2027</code>, так и чистый ID <code>2026</code>.</li>
                      <li><b>Защита от повторной подачи (дубликатов)</b>: перед отправкой заявки программа проверяет через <code className="text-purple-300">GET /api/rest/order</code> наличие активных заявок со статусами <code className="text-emerald-400">initial</code>, <code className="text-emerald-400">approve</code>, <code className="text-emerald-400">study</code> для данного ребенка и программы. Если заявка уже есть, создание пропускается, а в таблице фиксируется актуальный статус!</li>
                      <li>Опционально отметьте чекбоксы: <code className="text-purple-300">create_certificate: true</code> (создавать сертификат) и <code className="text-purple-300">use_certificate</code>.</li>
                      <li>Нажмите <b>«▶ Запустить зачисление детей на программу из таблицы»</b>. После успешной отправки заявки в 3-й столбец <code className="text-purple-300">programm_list.xlsx</code> записывается флаг <code className="text-cyan-300">Зачислен (Заявка #...)</code>.</li>
                    </ul>
                  </li>
                  <li>
                    <b className="text-sky-300">Вкладка 4 (Подтверждение и Отметка об обучении с защитой от чужих заявок):</b>
                    <ul className="list-circle list-inside pl-4 pt-1 space-y-1 text-slate-400">
                      <li>Использует отдельную таблицу <code className="text-sky-400">study_list.xlsx</code> (2 столбца: <b className="text-slate-300">ФИО</b> и <b className="text-slate-300">Флаг зачисления</b>).</li>
                      <li><b>Сравнение по Фамилии и Имени</b>: сайт Навигатора выдаёт в заявке Фамилию и Имя, а в таблице может быть полное ФИО или ФИ. Программа извлекает Фамилию и Имя с сайта и из таблицы, сравнивая только их (отчество игнорируется).</li>
                      <li><b>Защита от ошибочного формата с инициалами</b>: если в Excel фамилия указана с инициалами (например, <code>Иванов И.И.</code> или <code>Иванов И.</code>), программа отмечает такую запись во 2-м столбце: <code className="text-amber-400">⚠️ Неправильный формат (инициалы, укажите полное имя)</code> и блокирует её для предотвращения ошибок зачисления однофамильцев.</li>
                      <li><b>Защита от чужих заявок</b>: если родитель подал заявку на сайте на другую программу, которой нет в вашем бумажном заявлении и файле <code className="text-sky-400">study_list.xlsx</code>, программа пропустит такую заявку и не будет её подтверждать и зачислять.</li>
                      <li><b>Авто-флагирование</b>: при успешном подтверждении заявки флаг переключается на <code className="text-emerald-400">Подтвержден (approve)</code>, а при зачислении на обучение по приказу — на <code className="text-emerald-400">Зачислен (приказ №... от ...)</code>.</li>
                    </ul>
                  </li>
                </ul>
              </div>

              <div className="bg-slate-950 border border-slate-800 rounded-xl p-4 space-y-3 text-xs text-slate-300">
                <div className="flex items-center justify-between">
                  <div className="font-semibold text-slate-200 flex items-center gap-2">
                    <Zap className="w-4 h-4 text-amber-400" />
                    Шаг 4. Сборка в автономный EXE (запуск без Python):
                  </div>
                  <button
                    onClick={handleDownloadBat}
                    className="px-2.5 py-1 bg-amber-500/10 hover:bg-amber-500/20 text-amber-300 border border-amber-500/30 rounded flex items-center gap-1 text-[11px] font-medium"
                  >
                    <Download className="w-3 h-3" /> Скачать build_exe.bat
                  </button>
                </div>
                <p className="text-slate-400 leading-relaxed">
                  Чтобы конечный пользователь мог запускать программу обычным двойным кликом (без установки Python и библиотек), скрипт компилируется через <b>PyInstaller</b> в один <code className="text-slate-200">NavigatorApp.exe</code>.
                </p>
                <div className="bg-slate-900 border border-slate-800 rounded-lg p-3 space-y-1.5 font-mono text-[11px] text-emerald-400">
                  <div className="text-slate-500"># Быстрая сборка в одну команду:</div>
                  <div>pip install pyinstaller requests openpyxl</div>
                  <div>pyinstaller --noconsole --onefile --name "NavigatorApp" navigator_app.py</div>
                </div>
                <p className="text-slate-500 text-[11px]">
                  Готовый файл будет в папке <code className="text-slate-300">dist/NavigatorApp.exe</code>. В репозиторий также добавлен рабочий процесс <b>GitHub Actions</b> (<code>.github/workflows/build-exe.yml</code>), который автоматически создаёт сборку при публикации релиза.
                </p>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Modal: Проверка обновлений NavigatorApp */}
      {showUpdateModal && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl max-w-lg w-full p-6 shadow-2xl space-y-4">
            <div className="flex items-start justify-between">
              <div className="flex items-center gap-3">
                <div className="text-3xl">🚀</div>
                <div>
                  <h3 className="text-base font-bold text-slate-100">
                    Доступно обновление NavigatorApp
                  </h3>
                  <p className="text-xs text-slate-400">
                    Установлена: <span className="text-slate-300 font-mono">v2.1.0</span> | Репозиторий: <span className="text-blue-400 font-mono">Romosol/Navigator-Tools-REST-API-</span>
                  </p>
                </div>
              </div>
              <button
                onClick={() => setShowUpdateModal(false)}
                className="text-slate-400 hover:text-slate-200 text-lg leading-none cursor-pointer"
              >
                ✕
              </button>
            </div>

            <div className="bg-slate-950 border border-slate-800 rounded-xl p-3.5 space-y-2 text-xs">
              <div className="flex items-center justify-between text-slate-300 font-semibold">
                <span>Описание релиза (GitHub Releases)</span>
                <span className="text-emerald-400 text-[11px]">Latest Release</span>
              </div>
              <p className="text-slate-400 leading-relaxed">
                Свежий билд <code className="text-slate-200 font-mono">NavigatorApp.exe</code> доступен для прямой загрузки на GitHub.
                Сборка скомпилирована через GitHub Actions и содержит все свежие доработки.
              </p>
              <div className="pt-2 border-t border-slate-800 text-[11px] text-slate-400">
                ✅ Автономный запуск без Python<br />
                ✅ Встроенная проверка целостности и синтаксиса
              </div>
            </div>

            <div className="flex items-center justify-between pt-2">
              <label className="flex items-center gap-2 text-xs text-slate-400 cursor-pointer">
                <input type="checkbox" defaultChecked className="rounded border-slate-700 text-blue-600 focus:ring-blue-500" />
                <span>Проверять при каждом запуске</span>
              </label>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => setShowUpdateModal(false)}
                  className="px-3 py-1.5 rounded-lg border border-slate-700 hover:bg-slate-800 text-slate-300 text-xs font-medium transition-colors cursor-pointer"
                >
                  Напомнить позже
                </button>
                <a
                  href="https://github.com/Romosol/Navigator-Tools-REST-API-/releases/latest"
                  target="_blank"
                  rel="noreferrer"
                  className="px-3.5 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold flex items-center gap-1.5 transition-all shadow-md shadow-blue-600/20"
                >
                  <Download className="w-3.5 h-3.5" />
                  Скачать на GitHub
                </a>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Modal: Создать баг-репорт (GitHub Issues) */}
      {showBugModal && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl max-w-xl w-full p-6 shadow-2xl space-y-4 max-h-[90vh] overflow-y-auto">
            <div className="flex items-start justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-red-500/10 border border-red-500/30 flex items-center justify-center text-red-400">
                  <Bug className="w-6 h-6" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-100">
                    Создать баг-репорт (GitHub Issues)
                  </h3>
                  <p className="text-xs text-slate-400">
                    Отправка тикета в репозиторий <span className="text-red-400 font-mono">Romosol/Navigator-Tools-REST-API-</span>
                  </p>
                </div>
              </div>
              <button
                onClick={() => setShowBugModal(false)}
                className="text-slate-400 hover:text-slate-200 text-lg leading-none cursor-pointer"
              >
                ✕
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div>
                <label className="block text-slate-300 font-medium mb-1">
                  Тема ошибки (Заголовок):
                </label>
                <input
                  type="text"
                  value={bugTitle}
                  onChange={(e) => setBugTitle(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-slate-200 text-xs focus:border-red-500 focus:outline-none"
                  placeholder="Краткая суть проблемы..."
                />
              </div>

              <div>
                <label className="block text-slate-300 font-medium mb-1">
                  Раздел программы:
                </label>
                <select
                  value={bugSection}
                  onChange={(e) => setBugSection(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg px-3 py-2 text-slate-200 text-xs focus:border-red-500 focus:outline-none"
                >
                  <option>Вкладка 1: Пакетная запись на мероприятие (event_list.xlsx)</option>
                  <option>Вкладка 2: Подтверждение и участие в мероприятии</option>
                  <option>Вкладка 3: Зачисление на учебную программу (programm_list.xlsx)</option>
                  <option>Вкладка 4: Сверка со study_list, подтверждение и обучение</option>
                  <option>Экран авторизации / Вход в систему</option>
                  <option>Проверка обновлений или настройки</option>
                  <option>Другое / Общая ошибка приложения</option>
                </select>
              </div>

              <div>
                <label className="block text-slate-300 font-medium mb-1">
                  Что произошло и шаги для воспроизведения:
                </label>
                <textarea
                  rows={4}
                  value={bugDesc}
                  onChange={(e) => setBugDesc(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-700 rounded-lg p-2.5 text-slate-200 text-xs font-mono focus:border-red-500 focus:outline-none resize-none"
                />
              </div>

              <div className="bg-slate-950 border border-slate-800 rounded-xl p-3 space-y-2">
                <label className="flex items-center gap-2 text-slate-300 cursor-pointer select-none">
                  <input
                    type="checkbox"
                    checked={bugDiag}
                    onChange={(e) => setBugDiag(e.target.checked)}
                    className="rounded border-slate-700 text-red-600 focus:ring-red-500"
                  />
                  <span className="font-semibold text-slate-200">Прикрепить диагностику и системный журнал</span>
                </label>
                {bugDiag && (
                  <div className="text-[11px] text-slate-400 pl-5 font-mono space-y-0.5">
                    <div>• Версия: v2.1.0 (Windows / PyInstaller)</div>
                    <div>• Журнал: logs/results_log.txt (последние 15 строк с маскированием паролей)</div>
                  </div>
                )}
              </div>
            </div>

            <div className="flex items-center justify-between pt-2 border-t border-slate-800">
              <span className="text-[11px] text-slate-400">
                {bugCopied ? <span className="text-emerald-400 font-semibold">✓ Текст скопирован в буфер!</span> : 'Откроется форма GitHub с заполненными полями'}
              </span>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => {
                    const text = `# [BUG] ${bugTitle}\n\n### Описание проблемы\n${bugDesc}\n\n### Раздел программы\n- ${bugSection}\n\n### Диагностика\n- Версия: v2.1.0\n- ОС: Windows\n`;
                    navigator.clipboard.writeText(text);
                    setBugCopied(true);
                    setTimeout(() => setBugCopied(false), 2000);
                  }}
                  className="px-3 py-1.5 rounded-lg border border-slate-700 hover:bg-slate-800 text-slate-300 text-xs font-medium transition-colors cursor-pointer"
                >
                  📋 Скопировать текст
                </button>
                <a
                  href={`https://github.com/Romosol/Navigator-Tools-REST-API-/issues/new?title=${encodeURIComponent(`[BUG] ${bugTitle}`)}&body=${encodeURIComponent(`### Описание проблемы\n${bugDesc}\n\n### Раздел программы\n- ${bugSection}\n\n### Окружение\n- Версия NavigatorApp: v2.1.0\n- Способ запуска: NavigatorApp.exe\n`)}&labels=bug`}
                  target="_blank"
                  rel="noreferrer"
                  className="px-3.5 py-1.5 rounded-lg bg-red-600 hover:bg-red-500 text-white text-xs font-semibold flex items-center gap-1.5 transition-all shadow-md shadow-red-600/20"
                >
                  <Bug className="w-3.5 h-3.5" />
                  Отправить в GitHub Issues
                </a>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
