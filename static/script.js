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
  const adminButton = document.getElementById('admin-btn');
  const adminModal = document.getElementById('admin-modal');
  const adminModalClose = document.getElementById('admin-modal-close');
  const adminStatsContent = document.getElementById('admin-stats-content');
  const jurisdictionGroups = [
    document.getElementById('analysis-jurisdiction-group'),
    document.getElementById('generation-jurisdiction-group')
  ];
  let checksCount = Math.max(0, Number(localStorage.getItem('lexify_checks_count') ?? 3));
  let generatedDocumentBlob = null;
  let generatedDocumentFilename = 'Lexify_Document.docx';

  function selectedJurisdiction(containerId) {
    const container = document.getElementById(containerId);
    return container?.querySelector('.jurisdiction-btn.active')?.dataset.jurisdiction || 'Россия (РФ)';
  }

  function updateChecksCount() {
    ['checks-count', 'checks-count-analysis', 'checks-count-generation'].forEach((id) => {
      const element = document.getElementById(id);
      if (element) element.textContent = String(checksCount);
    });
  }

  function consumeCheck() {
    checksCount = Math.max(0, checksCount - 1);
    localStorage.setItem('lexify_checks_count', String(checksCount));
    updateChecksCount();
  }

  updateChecksCount();
  jurisdictionGroups.filter(Boolean).forEach((group) => {
    group.querySelectorAll('.jurisdiction-btn').forEach((button) => {
      button.addEventListener('click', () => {
        group.querySelectorAll('.jurisdiction-btn').forEach((item) => item.classList.remove('active'));
        button.classList.add('active');
      });
    });
  });

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
      consumeCheck();
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
      const jurisdiction = selectedJurisdiction('generation-jurisdiction-group');

      btnRunGenerate.disabled = true;
      try {
        const response = await fetch('/api/generate-document', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            doc_type: docType,
            parameters,
            init_data: '',
            jurisdiction
          })
        });

        if (!response.ok) {
          throw new Error(`Ошибка генерации документа: ${response.status}`);
        }

        const blob = await response.blob();
        const disposition = response.headers.get('Content-Disposition') || '';
        const filenameMatch = disposition.match(/filename=([^;]+)/i);
        const filename = filenameMatch ? filenameMatch[1].replace(/"/g, '') : 'Lexify_Document.docx';
        generatedDocumentBlob = blob;
        generatedDocumentFilename = filename;

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
        consumeCheck();
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
    return fetch('/api/generate-report', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        doc_type: filename,
        parameters: text,
        init_data: '',
        jurisdiction: selectedJurisdiction('analysis-jurisdiction-group')
      })
    }).then((response) => {
      if (!response.ok) throw new Error(`Ошибка формирования файла: ${response.status}`);
      return response.blob();
    }).then((blob) => {
      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = filename === 'protocol'
        ? 'Protocol_Raznoglasiy.docx'
        : 'Lexify_Result_Analysis.docx';
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(downloadUrl);
    }).catch((error) => console.error(error));
  }

  if (btnDownloadConclusion) {
    btnDownloadConclusion.addEventListener('click', () => {
      downloadReport('analysis',
        'Экспертное заключение Lexify. Требуется проведение детального правового анализа представленного договора.');
    });
  }
  if (btnDownloadProtocol) {
    btnDownloadProtocol.addEventListener('click', () => {
      downloadReport('protocol',
        'Перечень спорных условий по результатам правовой проверки и предложенные редакции.');
    });
  }
  if (btnDownloadGenerated) {
    btnDownloadGenerated.addEventListener('click', () => {
      if (!generatedDocumentBlob) {
        btnRunGenerate.click();
        return;
      }
      const downloadUrl = URL.createObjectURL(generatedDocumentBlob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = generatedDocumentFilename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(downloadUrl);
    });
  }

  if (adminButton && adminModal && adminModalClose && adminStatsContent) {
    adminButton.addEventListener('click', async () => {
      adminModal.classList.remove('hidden');
      adminStatsContent.textContent = 'Загрузка статистики...';
      try {
        const response = await fetch('/admin/stats/0');
        if (!response.ok) throw new Error(`Ошибка загрузки статистики: ${response.status}`);
        const stats = await response.json();
        adminStatsContent.innerHTML = `
          <div class="admin-stat"><strong>${stats.total_users}</strong>Пользователей</div>
          <div class="admin-stat"><strong>${stats.total_analyses}</strong>Анализов договоров</div>
          <div class="admin-stat"><strong>${stats.total_generations}</strong>Сгенерированных документов</div>
          <div class="admin-stat"><strong>${stats.total_payments}</strong>Платёжных операций</div>
        `;
      } catch (error) {
        adminStatsContent.textContent = 'Не удалось загрузить статистику админ-панели.';
        console.error(error);
      }
    });

    const closeAdminModal = () => adminModal.classList.add('hidden');
    adminModalClose.addEventListener('click', closeAdminModal);
    adminModal.addEventListener('click', (event) => {
      if (event.target === adminModal) closeAdminModal();
    });
  }
});
