const form = document.querySelector('#upload-form');
const sourceFile = document.querySelector('#source-file');
const refreshButton = document.querySelector('#refresh-button');
const submitButton = document.querySelector('#submit-button');
const statusMessage = document.querySelector('#status');
const results = document.querySelector('#results');

sourceFile.addEventListener('change', () => {
  results.replaceChildren();
  statusMessage.hidden = true;
  submitButton.disabled = !sourceFile.value;
});

refreshButton.addEventListener('click', loadFiles);

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  if (!sourceFile.value) return;

  submitButton.disabled = true;
  refreshButton.disabled = true;
  submitButton.textContent = 'Przetwarzanie…';
  results.replaceChildren();
  statusMessage.className = 'status';
  statusMessage.textContent = 'Odczytujemy wybrany plik i uzupełniamy szablon.';
  statusMessage.hidden = false;

  try {
    const response = await fetch('/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: sourceFile.value }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || 'Nie udało się wygenerować pliku.');

    statusMessage.classList.add('success');
    statusMessage.textContent = `Gotowe. Pliki zapisano w folderze exports (${payload.files.length}).`;
    for (const file of payload.files) {
      const card = document.createElement('article');
      card.className = 'result-card';

      const details = document.createElement('div');
      const title = document.createElement('strong');
      title.textContent = `Ewidencja ${file.month}`;
      const summary = document.createElement('span');
      summary.textContent = `${file.count} transakcji · suma ${formatAmount(file.total)} zł`;
      details.append(title, summary);

      const link = document.createElement('a');
      link.href = file.url;
      link.download = file.filename;
      link.textContent = 'Pobierz';
      card.append(details, link);
      results.append(card);
    }
  } catch (error) {
    statusMessage.classList.add('error');
    statusMessage.textContent = error.message || 'Nie udało się wygenerować pliku.';
  } finally {
    submitButton.disabled = false;
    refreshButton.disabled = false;
    submitButton.innerHTML = 'Utwórz ewidencję <span aria-hidden="true">→</span>';
  }
});

async function loadFiles() {
  const previousSelection = sourceFile.value;
  refreshButton.disabled = true;
  submitButton.disabled = true;
  sourceFile.disabled = true;
  sourceFile.replaceChildren(new Option('Wczytywanie plików…', ''));
  statusMessage.hidden = true;
  results.replaceChildren();

  try {
    const response = await fetch('/api/files', { cache: 'no-store' });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || 'Nie udało się odczytać folderu dane.');

    sourceFile.replaceChildren();
    if (payload.files.length === 0) {
      sourceFile.add(new Option('Brak plików transakcji w folderze dane', ''));
      statusMessage.className = 'status';
      statusMessage.textContent = 'Umieść plik „transakcje <data>.xlsx” w folderze dane i odśwież listę.';
      statusMessage.hidden = false;
      return;
    }

    sourceFile.add(new Option('Wybierz plik transakcji…', ''));
    for (const file of payload.files) {
      sourceFile.add(new Option(file.filename, file.filename));
    }
    if (payload.files.some((file) => file.filename === previousSelection)) {
      sourceFile.value = previousSelection;
    }
  } catch (error) {
    sourceFile.replaceChildren(new Option('Nie udało się wczytać listy plików', ''));
    statusMessage.className = 'status error';
    statusMessage.textContent = error.message || 'Nie udało się odczytać folderu dane.';
    statusMessage.hidden = false;
  } finally {
    sourceFile.disabled = false;
    refreshButton.disabled = false;
    submitButton.disabled = !sourceFile.value;
  }
}

function formatAmount(amount) {
  return new Intl.NumberFormat('pl-PL', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(amount);
}

loadFiles();
