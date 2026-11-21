/* ========================================== */
/* LEXIFY: ЛОГИКА И ПЕРЕКЛЮЧЕНИЕ ОКОН          */
/* ========================================== */

document.addEventListener('DOMContentLoaded', () => {
  const windowMain = document.getElementById('window-main');
  const windowAnalysis = document.getElementById('window-analysis');
  const windowGenerate = document.getElementById('window-generate');

  const btnGotoAnalysis = document.getElementById('btn-goto-analysis');
  const btnGotoGenerate = document.getElementById('btn-goto-generate');

  const btnBackFromAnalysis = document.getElementById('btn-back-from-analysis');
  const btnBackFromGenerate = document.getElementById('btn-back-from-generate');

  const modeText = document.getElementById('mode-text');
  const modeFile = document.getElementById('mode-file');
  const btnRunAnalysis = document.getElementById('btn-run-analysis');
  const analysisResultBox = document.getElementById('analysis-result-box');
  const btnDownloadConclusion = document.getElementById('btn-download-conclusion');
  const btnDownloadProtocol = document.getElementById('btn-download-protocol');

  const typeButtons = document.querySelectorAll('#window-generate .type-btn');
  const btnRunGenerate = document.getElementById('btn-run-generate');
  const genResultBox = document.getElementById('gen-result-box');
  const btnDownloadGenerated = document.getElementById('btn-download-generated');
  const genParamsInput = document.getElementById('gen-params-input');
  const genResultSummary = document.getElementById('gen-result-summary');

  function showWindow(windowToShow) {
    if (!windowToShow) return;
    windowMain.classList.add('hidden');
    windowAnalysis.classList.add('hidden');
    windowGenerate.classList.add('hidden');

    windowToShow.classList.remove('hidden');
    window.scrollTo(0, 0);
  }

  if (btnGotoAnalysis) {
    btnGotoAnalysis.addEventListener('click', () => showWindow(windowAnalysis));
  }
  if (btnGotoGenerate) {
    btnGotoGenerate.addEventListener('click', () => showWindow(windowGenerate));
  }

  if (btnBackFromAnalysis) {
    btnBackFromAnalysis.addEventListener('click', () => showWindow(windowMain));
  }
  if (btnBackFromGenerate) {
    btnBackFromGenerate.addEventListener('click', () => showWindow(windowMain));
  }

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

  if (btnRunAnalysis && analysisResultBox) {
    btnRunAnalysis.addEventListener('click', () => {
      analysisResultBox.classList.remove('hidden');
    });
  }

  typeButtons.forEach((btn) => {
    btn.addEventListener('click', () => {
      typeButtons.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      if (genResultBox) genResultBox.classList.add('hidden');
    });
  });

  if (btnRunGenerate && genResultBox) {
    btnRunGenerate.addEventListener('click', async () => {
      const activeType = document.querySelector('#window-generate .type-btn.active');
      const docType = activeType ? activeType.dataset.docType : 'contract';
      const parameters = genParamsInput ? genParamsInput.value.trim() : '';

      btnRunGenerate.disabled = true;
      try {
        const response = await fetch('/api/generate-document', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            doc_type: docType,
            parameters,
            init_data: ''
          })
        });

        if (!response.ok) {
          throw new Error(`Ошибка генерации документа: ${response.status}`);
        }

        const blob = await response.blob();
        const disposition = response.headers.get('Content-Disposition') || '';
        const filenameMatch = disposition.match(/filename=([^;]+)/i);
        const filename = filenameMatch ? filenameMatch[1].replace(/"/g, '') : 'Lexify_Document.docx';

        const downloadUrl = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = downloadUrl;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(downloadUrl);

        if (genResultSummary) {
          genResultSummary.textContent = 'Проект документа сформирован и передан на загрузку.';
        }
        genResultBox.classList.remove('hidden');
      } catch (error) {
        if (genResultSummary) {
          genResultSummary.textContent = 'Не удалось сформировать документ. Повторите попытку.';
        }
        genResultBox.classList.remove('hidden');
        console.error(error);
      } finally {
        btnRunGenerate.disabled = false;
      }
    });
  }

  function downloadReport(filename, text) {
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(downloadUrl);
  }

  if (btnDownloadConclusion) {
    btnDownloadConclusion.addEventListener('click', () => {
      downloadReport('Lexify_Expertnoe_zaklyuchenie.txt',
        'Экспертное заключение Lexify\n\nТребуется проведение детального правового анализа представленного договора.');
    });
  }
  if (btnDownloadProtocol) {
    btnDownloadProtocol.addEventListener('click', () => {
      downloadReport('Lexify_Protokol_raznoglasiy.txt',
        'Протокол разногласий Lexify\n\nПеречень спорных условий будет сформирован по результатам правовой проверки.');
    });
  }
  if (btnDownloadGenerated) {
    btnDownloadGenerated.addEventListener('click', () => {
      btnRunGenerate.click();
    });
  }
});
