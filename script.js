let slidesData = [];

Office.onReady(() => {
    console.log("Office ready");
    loadCatalog();
});

async function loadCatalog() {
    const grid = document.getElementById('slidesGrid');
    grid.innerHTML = '<div class="loading">Загрузка каталога...</div>';
    try {
        const response = await fetch('catalog.json');
        if (!response.ok) throw new Error(`HTTP ${response.status} - ${response.statusText}`);
        slidesData = await response.json();
        if (!slidesData.length) throw new Error('Каталог пуст');
        renderSlides(slidesData);
    } catch (err) {
        grid.innerHTML = `<div class="loading" style="color:red;">Ошибка загрузки: ${err.message}</div>`;
    }
}

function renderSlides(slides) {
    const grid = document.getElementById('slidesGrid');
    if (!slides.length) {
        grid.innerHTML = '<div class="loading">Нет слайдов</div>';
        return;
    }
    grid.innerHTML = slides.map(slide => `
        <div class="slide-card">
            <div class="preview">
                ${slide.previewBase64 ? `<img src="${slide.previewBase64}" alt="${slide.title}">` : '<div class="placeholder">📄 Нет превью</div>'}
            </div>
            <div class="card-info">
                <div class="card-title">${escapeHtml(slide.title)}</div>
                <div class="card-category">${slide.category}</div>
                <div class="card-tags">${slide.tags.join(', ')}</div>
                <button class="insert-btn" data-id="${slide.id}">➕ Insert slide</button>
            </div>
        </div>
    `).join('');

    document.querySelectorAll('.insert-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const slideId = btn.getAttribute('data-id');
            const slide = slidesData.find(s => s.id == slideId);
            if (slide && slide.slideBase64) {
                insertSlide(slide.slideBase64, slide.title);
            } else {
                alert('Ошибка: нет данных слайда');
            }
        });
    });
}

function insertSlide(base64Data, title) {
    if (!Office.context || !Office.context.document) {
        alert('Ошибка: Office API не инициализирован');
        return;
    }
    Office.context.document.insertSlidesFromBase64(base64Data, {
        format: Office.FileType.Compressed
    }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            alert('Ошибка вставки: ' + result.error.message);
        } else {
            alert('Слайд "' + title + '" вставлен!');
        }
    });
}

function escapeHtml(str) {
    return str.replace(/[&<>]/g, function(m) {
        if (m === '&') return '&amp;';
        if (m === '<') return '&lt;';
        if (m === '>') return '&gt;';
        return m;
    });
}
