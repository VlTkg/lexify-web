/* ========================================== */
/* LEXIFY: ЛОГИКА И ПЕРЕКЛЮЧЕНИЕ ОКОН          */
/* ========================================== */

document.addEventListener('DOMContentLoaded', () => {
  // Элементы окон
  const windowMain = document.getElementById('window-main');
  const windowAnalysis = document.getElementById('window-analysis');
  const windowGenerate = document.getElementById('window-generate');

  // Кнопки навигации с главного меню
  const btnGotoAnalysis = document.getElementById('btn-goto-analysis');
  const btnGotoGenerate = document.getElementById('btn-goto-generate');

  // Кнопки возврата
  const btnBackFromAnalysis = document.getElementById('btn-back-from-analysis');
  const btnBackFromGenerate = document.getElementById('btn-back-from-generate');

  // Элементы Окна 2 (Анализ)
  const modeText = document.getElementById('mode-text');
  const modeFile = document.getElementById('mode-file');
  const btnRunAnalysis = document.getElementById('btn-run-analysis');
  const analysisResultBox = document.getElementById('analysis-result-box');
  const btnDownloadConclusion = document.getElementById('btn-download-conclusion');
  const btnDownloadProtocol = document.getElementById('btn-download-protocol');

  // Элементы Окна 3 (Генерация)
  const typeButtons = document.querySelectorAll('#window-generate .type-btn');
  const btnRunGenerate = document.getElementById('btn-run-generate');
  const genResultBox = document.getElementById('gen-result-box');
  const btnDownloadGenerated = document.getElementById('btn-download-generated');

  // Функция переключения видимости окон
  function showWindow(windowToShow) {
    windowMain.classList.add('hidden');
    windowAnalysis.classList.add('hidden');
    windowGenerate.classList.add('hidden');

    windowToShow.classList.remove('hidden');
    window.scrollTo(0, 0);
  }

  // Навигация из главного меню
  if (btnGotoAnalysis) {
    btnGotoAnalysis.addEventListener('click', () => showWindow(windowAnalysis));
  }
  if (btnGotoGenerate) {
    btnGotoGenerate.addEventListener('click', () => showWindow(windowGenerate));
  }

  // Возврат в главное меню
  if (btnBackFromAnalysis) {
    btnBackFromAnalysis.addEventListener('click', () => showWindow(windowMain));
  }
  if (btnBackFromGenerate) {
    btnBackFromGenerate.addEventListener('click', () => showWindow(windowMain));
  }

  // Переключение режимов в Окне 2 (Текст / Файл)
  if (modeText && modeFile) {
    modeText.addEventListener('click', () => {
      modeText.classList.add('active');
      modeFile.classList.remove('active');
    });
    modeFile.addEventListener('click', () => {
      modeFile.classList.add('active');
      modeText.classList.remove('active');
    });
  }

  // Запуск анализа в Окне 2 (появление блока результатов с двумя кнопками)
  if (btnRunAnalysis && analysisResultBox) {
    btnRunAnalysis.addEventListener('click', () => {
      analysisResultBox.classList.remove('hidden');
    });
  }

  // Переключение типов документов в Окне 3
  typeButtons.forEach((btn) => {
    btn.addEventListener('click', () => {
      typeButtons.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      if (genResultBox) genResultBox.classList.add('hidden');
    });
  });

  // Запуск генерации в Окне 3 (появление результата и кнопки скачивания)
  if (btnRunGenerate && genResultBox) {
    btnRunGenerate.addEventListener('click', () => {
      genResultBox.classList.remove('hidden');
    });
  }

  // Заглушки для скачивания файлов (.docx)
  if (btnDownloadConclusion) {
    btnDownloadConclusion.addEventListener('click', () => {
      alert('Загрузка файла: Expert_Conclusion.docx');
    });
  }
  if (btnDownloadProtocol) {
    btnDownloadProtocol.addEventListener('click', () => {
      alert('Загрузка файла: Protocol_of_Disagreements.docx');
    });
  }
  if (btnDownloadGenerated) {
    btnDownloadGenerated.addEventListener('click', () => {
      alert('Загрузка файла: Generated_Document.docx');
    });
  }
});
