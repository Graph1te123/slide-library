let slidesData = [];

Office.onReady(() => {
    console.log("Office ready");
    loadCatalog();
});

async function loadCatalog() {
    try {
        const response = await fetch('catalog.json');
        if (!response.ok) throw new Error('Failed to load catalog');
        slidesData = await response.json();
        renderSlides(slidesData);
    } catch (error) {
        document.getElementById('slidesGrid').innerHTML = `<div class="loading">Error loading catalog: ${error.message}</div>`;
    }
}

function renderSlides(slides) {
    const grid = document.getElementById('slidesGrid');
    if (!slides.length) {
        grid.innerHTML = '<div class="loading">No slides found</div>';
        return;
    }
    grid.innerHTML = slides.map(slide => `
        <div class="slide-card" data-id="${slide.id}">
            <div class="preview">
                ${slide.previewBase64 ? 
                    `<img src="${slide.previewBase64}" alt="${slide.title}">` : 
                    `<div class="placeholder">📄 No preview</div>`
                }
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
            e.stopPropagation();
            const slideId = btn.getAttribute('data-id');
            const slide = slidesData.find(s => s.id == slideId);
            if (slide && slide.slideBase64) {
                insertSlide(slide.slideBase64, slide.title);
            } else {
                console.error("No base64 data for slide", slideId);
            }
        });
    });
}

function insertSlide(base64Data, title) {
    Office.context.document.insertSlidesFromBase64(base64Data, {
        format: Office.FileType.Compressed
    }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Insert failed", result.error.message);
        } else {
            console.log(`Slide "${title}" inserted`);
        }
    });
}

document.getElementById('searchInput').addEventListener('input', filterSlides);
document.getElementById('categoryFilter').addEventListener('change', filterSlides);

function filterSlides() {
    const query = document.getElementById('searchInput').value.toLowerCase();
    const category = document.getElementById('categoryFilter').value;
    const filtered = slidesData.filter(slide => {
        const matchesSearch = slide.title.toLowerCase().includes(query) || 
                              slide.tags.some(tag => tag.toLowerCase().includes(query));
        const matchesCategory = category === 'all' || slide.category === category;
        return matchesSearch && matchesCategory;
    });
    renderSlides(filtered);
}

function escapeHtml(str) {
    return str.replace(/[&<>]/g, function(m) {
        if (m === '&') return '&amp;';
        if (m === '<') return '&lt;';
        if (m === '>') return '&gt;';
        return m;
    });
}
