const form = document.getElementById('reportForm');
const statusEl = document.getElementById('status');
const submitBtn = document.getElementById('submitBtn');

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.style.color = isError ? '#fca5a5' : '#67e8f9';
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();

  submitBtn.disabled = true;
  setStatus('Generuji elaborát...');

  try {
    const formData = new FormData(form);
    const apiKey = (formData.get('api_key') || '').toString().trim();

    if (!apiKey) {
      throw new Error('API key je povinný.');
    }

    // checkbox -> backend očekává bool hodnotu
    if (!formData.get('is_handwritten')) {
      formData.set('is_handwritten', 'false');
    }

    const response = await fetch('/api/generate', {
      method: 'POST',
      body: formData,
    });

    if (!response.ok) {
      let errorText = `Chyba ${response.status}`;
      try {
        const err = await response.json();
        errorText = err.error || JSON.stringify(err);
      } catch (_) {
        errorText = await response.text();
      }
      throw new Error(errorText);
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'laboratorni_protokol.docx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    const qualityStatus = response.headers.get('X-Quality-Status') || 'PASS';
    setStatus(`Hotovo. Soubor stažen. Quality gate: ${qualityStatus}`);
  } catch (error) {
    setStatus(`Generování selhalo: ${error.message}`, true);
  } finally {
    submitBtn.disabled = false;
  }
});
