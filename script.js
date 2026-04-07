let slidesData = [];

Office.onReady(() => {
    console.log("Office ready");
    loadCatalog();

    document.getElementById('searchInput').addEventListener('input', filterSlides);
    document.getElementById('categoryFilter').addEventListener('change', filterSlides);
});

async function loadCatalog() {
    const grid = document.getElementById('slidesGrid');
    grid.innerHTML = '<div class="loading">Loading...</div>';

    try {
        const response = await fetch('catalog.json');
        slidesData = await response.json();
        renderSlides(slidesData);
    } catch (err) {
        grid.innerHTML = `<div class="loading" style="color:red;">Error: ${err.message}</div>`;
    }
}

function renderSlides(slides) {
    const grid = document.getElementById('slidesGrid');

    if (!slides.length) {
        grid.innerHTML = '<div class="loading">No slides found</div>';
        return;
    }

    grid.innerHTML = slides.map(slide => `
    <div class="slide-card">
        <div class="preview">
            ${slide.previewBase64
                ? `<img src="${slide.previewBase64}">`
                : '<div class="placeholder">No preview</div>'}
        </div>
        <div class="card-info">
            <div class="card-title">${slide.title}</div>
            <div class="card-category">${slide.category}</div>
            <div class="card-tags">${slide.tags.join(', ')}</div>
            <button class="insert-btn" data-id="${slide.id}">Insert</button>
        </div>
    </div>
    `).join('');

    document.querySelectorAll('.insert-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const slideId = btn.dataset.id;
            const slide = slidesData.find(s => s.id == slideId);
            insertSlide(slide);
        });
    });
}

function filterSlides() {
    const query = document.getElementById('searchInput').value.toLowerCase();
    const category = document.getElementById('categoryFilter').value;

    const filtered = slidesData.filter(slide => {
        const matchQuery =
            slide.title.toLowerCase().includes(query) ||
            slide.tags.some(tag => tag.toLowerCase().includes(query));

        const matchCategory =
            category === 'all' || slide.category === category;

        return matchQuery && matchCategory;
    });

    renderSlides(filtered);
}

function insertSlide(slide) {

    if (!slide || !slide.slideBase64) {
        alert("No slide data");
        return;
    }

    if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.2')) {

        const link = document.createElement('a');
        link.href = "data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64," + slide.slideBase64;
        link.download = slide.title + ".pptx";
        link.click();

        alert("PowerPoint Online doesn't support direct insert. File downloaded.");
        return;
    }


    Office.context.document.insertSlidesFromBase64(slide.slideBase64, {
        format: Office.FileType.Compressed
    }, (result) => {

        if (result.status === Office.AsyncResultStatus.Failed) {
            alert("Error: " + result.error.message);
        } else {
            console.log("Slide inserted");
        }

    });
}
