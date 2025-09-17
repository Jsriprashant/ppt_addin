// Wait for Office to initialize
Office.onReady((info) => {
    if (info.host === Office.HostType.PowerPoint) {
        console.log('PowerPoint add-in loaded successfully');
        document.addEventListener('DOMContentLoaded', initializeAddin);
        // If DOM is already loaded
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', initializeAddin);
        } else {
            initializeAddin();
        }
    }
});

const apiBase = window.location.origin + '/api';

function setStatus(msg) {
    const statusElement = document.getElementById('status');
    if (statusElement) {
        statusElement.textContent = msg;
        console.log('Status:', msg);
    }
}

async function getSelectedText() {
    return new Promise((resolve, reject) => {
        try {
            PowerPoint.run(async (context) => {
                // Get the selected shapes
                const selectedShapes = context.presentation.getSelectedShapes();
                const selectedSlides = context.presentation.getSelectedSlides();

                // Load the shapes to get their properties
                selectedShapes.load('items');
                await context.sync();

                let extractedText = '';

                if (selectedShapes.items.length > 0) {
                    // Extract text from selected shapes
                    for (let shape of selectedShapes.items) {
                        if (shape.textFrame) {
                            shape.textFrame.load('textRange');
                            await context.sync();
                            if (shape.textFrame.textRange && shape.textFrame.textRange.text) {
                                extractedText += shape.textFrame.textRange.text + '\n';
                            }
                        }
                    }
                } else if (selectedSlides.items && selectedSlides.items.length > 0) {
                    // If no shapes selected, get text from selected slide
                    selectedSlides.load('items');
                    await context.sync();

                    const slide = selectedSlides.items[0];
                    slide.shapes.load('items');
                    await context.sync();

                    for (let shape of slide.shapes.items) {
                        if (shape.textFrame) {
                            shape.textFrame.load('textRange');
                            await context.sync();
                            if (shape.textFrame.textRange && shape.textFrame.textRange.text) {
                                extractedText += shape.textFrame.textRange.text + '\n';
                            }
                        }
                    }
                }

                resolve(extractedText.trim());
            }).catch((error) => {
                console.error('Error getting selected text:', error);
                reject(error);
            });
        } catch (error) {
            console.error('Error in getSelectedText:', error);
            reject(error);
        }
    });
}

async function setSelectedText(text) {
    return new Promise((resolve, reject) => {
        try {
            PowerPoint.run(async (context) => {
                const selectedShapes = context.presentation.getSelectedShapes();
                selectedShapes.load('items');
                await context.sync();

                if (selectedShapes.items.length > 0) {
                    // Replace text in selected shapes
                    for (let shape of selectedShapes.items) {
                        if (shape.textFrame) {
                            shape.textFrame.textRange.text = text;
                        }
                    }
                } else {
                    // If no shape selected, add text to current slide
                    const slides = context.presentation.slides;
                    slides.load('items');
                    await context.sync();

                    if (slides.items.length > 0) {
                        const currentSlide = slides.items[0]; // Or get the active slide
                        const textBox = currentSlide.shapes.addTextBox(text);
                        textBox.left = 100;
                        textBox.top = 100;
                        textBox.height = 200;
                        textBox.width = 400;
                    }
                }

                await context.sync();
                resolve();
            }).catch((error) => {
                console.error('Error setting selected text:', error);
                reject(error);
            });
        } catch (error) {
            console.error('Error in setSelectedText:', error);
            reject(error);
        }
    });
}

async function readSelected() {
    try {
        setStatus('Reading selection...');
        const text = await getSelectedText();
        const textArea = document.getElementById('selectedText');
        if (textArea) {
            textArea.value = text;
        }
        setStatus(text ? 'Selected text loaded' : 'No text found in selection');
    } catch (err) {
        const errorMsg = 'Error reading selection: ' + (err.message || JSON.stringify(err));
        console.error(errorMsg);
        setStatus(errorMsg);
    }
}

async function saveSelected() {
    try {
        setStatus('Saving selection...');
        const text = await getSelectedText();

        if (!text.trim()) {
            setStatus('No text to save');
            return;
        }

        const response = await fetch(apiBase + '/texts', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ text })
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const data = await response.json();
        setStatus('Saved with ID: ' + data.id);
        await refreshList();
    } catch (err) {
        const errorMsg = 'Error saving: ' + (err.message || err);
        console.error(errorMsg);
        setStatus(errorMsg);
    }
}

async function loadById() {
    try {
        const idInput = document.getElementById('textId');
        const id = idInput ? idInput.value.trim() : '';

        if (!id) {
            setStatus('Please enter an ID');
            return;
        }

        setStatus('Loading ID ' + id + '...');
        const response = await fetch(apiBase + '/texts/' + encodeURIComponent(id));

        if (!response.ok) {
            setStatus('Text not found for ID: ' + id);
            return;
        }

        const data = await response.json();
        await setSelectedText(data.text);
        setStatus('Loaded text into selection');
    } catch (err) {
        const errorMsg = 'Error loading: ' + (err.message || err);
        console.error(errorMsg);
        setStatus(errorMsg);
    }
}

async function updateById() {
    try {
        const idInput = document.getElementById('textId');
        const id = idInput ? idInput.value.trim() : '';

        if (!id) {
            setStatus('Please enter an ID');
            return;
        }

        const text = await getSelectedText();
        setStatus('Updating ID ' + id + '...');

        const response = await fetch(apiBase + '/texts/' + encodeURIComponent(id), {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ text })
        });

        if (!response.ok) {
            setStatus('Update failed for ID: ' + id);
            return;
        }

        setStatus('Updated ID ' + id);
        await refreshList();
    } catch (err) {
        const errorMsg = 'Error updating: ' + (err.message || err);
        console.error(errorMsg);
        setStatus(errorMsg);
    }
}

async function deleteSelectedText() {
    try {
        setStatus('Clearing selected text...');
        await setSelectedText('');
        setStatus('Cleared selected text');
    } catch (err) {
        const errorMsg = 'Error deleting: ' + (err.message || err);
        console.error(errorMsg);
        setStatus(errorMsg);
    }
}

async function deleteById() {
    try {
        const idInput = document.getElementById('textId');
        const id = idInput ? idInput.value.trim() : '';

        if (!id) {
            setStatus('Please enter an ID to delete');
            return;
        }

        setStatus('Deleting ID ' + id + '...');
        const response = await fetch(apiBase + '/texts/' + encodeURIComponent(id), {
            method: 'DELETE'
        });

        if (!response.ok) {
            setStatus('Delete failed for ID: ' + id);
            return;
        }

        setStatus('Deleted ID ' + id);
        await refreshList();

        // Clear the ID input
        if (idInput) {
            idInput.value = '';
        }
    } catch (err) {
        const errorMsg = 'Error deleting: ' + (err.message || err);
        console.error(errorMsg);
        setStatus(errorMsg);
    }
}

async function refreshList() {
    try {
        const response = await fetch(apiBase + '/texts');

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const items = await response.json();
        const ul = document.getElementById('items');

        if (ul) {
            ul.innerHTML = '';

            if (items.length === 0) {
                const li = document.createElement('li');
                li.textContent = 'No saved texts';
                li.style.fontStyle = 'italic';
                ul.appendChild(li);
            } else {
                items.forEach(item => {
                    const li = document.createElement('li');
                    const preview = item.text.length > 50 ?
                        item.text.substring(0, 50) + '...' : item.text;
                    li.textContent = `${item.id}: ${preview}`;
                    li.style.cursor = 'pointer';
                    li.title = 'Click to load this text';

                    // Click to load functionality
                    li.onclick = () => {
                        const idInput = document.getElementById('textId');
                        if (idInput) {
                            idInput.value = item.id;
                        }
                        loadById();
                    };

                    ul.appendChild(li);
                });
            }
        }
    } catch (err) {
        const errorMsg = 'Error listing: ' + (err.message || err);
        console.error(errorMsg);
        setStatus(errorMsg);
    }
}

function initializeAddin() {
    console.log('Initializing add-in...');

    // Wire up buttons
    const btnRead = document.getElementById('btnRead');
    const btnSave = document.getElementById('btnSave');
    const btnLoad = document.getElementById('btnLoad');
    const btnUpdate = document.getElementById('btnUpdate');
    const btnDelete = document.getElementById('btnDelete');
    const btnDeleteById = document.getElementById('btnDeleteById');
    const btnRefresh = document.getElementById('btnRefresh');

    if (btnRead) btnRead.onclick = readSelected;
    if (btnSave) btnSave.onclick = saveSelected;
    if (btnLoad) btnLoad.onclick = loadById;
    if (btnUpdate) btnUpdate.onclick = updateById;
    if (btnDelete) btnDelete.onclick = deleteSelectedText;
    if (btnDeleteById) btnDeleteById.onclick = deleteById;
    if (btnRefresh) btnRefresh.onclick = refreshList;

    // Initial setup
    setStatus('Add-in loaded successfully');
    refreshList();

    console.log('Add-in initialization complete');
}