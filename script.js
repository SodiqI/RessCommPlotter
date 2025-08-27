
// Initialize the map with OpenStreetMap tiles
const map = L.map('map').setView([51.505, -0.09], 13);

// Add OpenStreetMap tile layer
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
}).addTo(map);

// Geodesic area calculation function (Shoelace formula with Earth's curvature)
function calculateGeodesicArea(latLngs) {
    if (latLngs.length < 3) return 0;
    
    const earthRadius = 6371000; // Earth's radius in meters
    let area = 0;
    
    for (let i = 0; i < latLngs.length; i++) {
        const j = (i + 1) % latLngs.length;
        const lat1 = latLngs[i].lat * Math.PI / 180;
        const lng1 = latLngs[i].lng * Math.PI / 180;
        const lat2 = latLngs[j].lat * Math.PI / 180;
        const lng2 = latLngs[j].lng * Math.PI / 180;
        
        area += (lng2 - lng1) * (2 + Math.sin(lat1) + Math.sin(lat2));
    }
    
    area = Math.abs(area * earthRadius * earthRadius / 2);
    return area;
}

// Initialize variables
let drawnItems = new L.FeatureGroup();
map.addLayer(drawnItems);

let polygons = [];
let currentPolygon = null;
let currentPoints = [];
let currentPolyline = null;
let nextId = 1;
let markers = [];
let selectedPolygonId = null;
let excelData = null;
let excelColumns = [];

// Undo/Redo system
let undoStack = [];
let redoStack = [];
let isEditMode = false;
let editingPolygonId = null;

// DOM Elements
const polygonIdInput = document.getElementById('polygon-id');
const polygonNameInput = document.getElementById('polygon-name');
const polygonCategorySelect = document.getElementById('polygon-category');
const latInput = document.getElementById('lat-input');
const lngInput = document.getElementById('lng-input');
const addCoordBtn = document.getElementById('add-coord-btn');
const bulkCoordsInput = document.getElementById('bulk-coords');
const addBulkBtn = document.getElementById('add-bulk-btn');
const plotBtn = document.getElementById('plot-btn');
const newBtn = document.getElementById('new-btn');
const deleteBtn = document.getElementById('delete-btn');
const refreshBtn = document.getElementById('refresh-btn');

// Add undo/redo functionality
function saveState() {
    const state = {
        polygons: JSON.parse(JSON.stringify(polygons.map(p => ({
            id: p.id,
            name: p.name,
            category: p.category,
            points: p.points,
            latLngs: p.latLngs,
            area: p.area,
            hectares: p.hectares,
            sqKm: p.sqKm,
            attributes: p.attributes
        })))),
        nextId: nextId
    };
    undoStack.push(state);
    if (undoStack.length > 20) { // Limit undo stack to 20 states
        undoStack.shift();
    }
    redoStack = []; // Clear redo stack when new action is performed
}

function undo() {
    if (undoStack.length === 0) {
        alert('Nothing to undo');
        return;
    }
    
    // Save current state to redo stack
    const currentState = {
        polygons: JSON.parse(JSON.stringify(polygons.map(p => ({
            id: p.id,
            name: p.name,
            category: p.category,
            points: p.points,
            latLngs: p.latLngs,
            area: p.area,
            hectares: p.hectares,
            sqKm: p.sqKm,
            attributes: p.attributes
        })))),
        nextId: nextId
    };
    redoStack.push(currentState);
    
    // Restore previous state
    const previousState = undoStack.pop();
    restoreState(previousState);
}

function redo() {
    if (redoStack.length === 0) {
        alert('Nothing to redo');
        return;
    }
    
    // Save current state to undo stack
    const currentState = {
        polygons: JSON.parse(JSON.stringify(polygons.map(p => ({
            id: p.id,
            name: p.name,
            category: p.category,
            points: p.points,
            latLngs: p.latLngs,
            area: p.area,
            hectares: p.hectares,
            sqKm: p.sqKm,
            attributes: p.attributes
        })))),
        nextId: nextId
    };
    undoStack.push(currentState);
    
    // Restore next state
    const nextState = redoStack.pop();
    restoreState(nextState);
}

function restoreState(state) {
    // Clear existing polygons from map
    polygons.forEach(p => {
        if (p.layer && p.layer._map) {
            map.removeLayer(p.layer);
        }
    });
    
    // Restore polygons
    polygons = [];
    nextId = state.nextId;
    
    state.polygons.forEach(polygonData => {
        // Recreate polygon layer
        const polygon = L.polygon([...polygonData.latLngs, polygonData.latLngs[0]], {
            color: getColorForCategory(polygonData.category),
            fillOpacity: 0.5,
            weight: 2
        }).addTo(map);
        
        polygon.bindPopup(`
            <strong>${polygonData.name}</strong><br>
            ID: ${polygonData.id}<br>
            Category: ${polygonData.category}<br>
            Points: ${polygonData.points.length}<br>
            Area: ${polygonData.area.toFixed(2)} m²<br>
            Hectares: ${polygonData.hectares.toFixed(4)} ha<br>
            Sq Km: ${polygonData.sqKm.toFixed(6)} sq km
        `);
        
        polygonData.layer = polygon;
        polygons.push(polygonData);
    });
    
    selectedPolygonId = null;
    updateShapesList();
    addNewPolygon();
}
const exportKmzBtn = document.getElementById('export-kmz');
const exportShpBtn = document.getElementById('export-shp');
const coordinatesDisplay = document.getElementById('coordinates-list');
const areaDisplay = document.getElementById('area-display');
const shapesContainer = document.getElementById('shapes-container');

// Set up event listeners
addCoordBtn.addEventListener('click', addCoordinateManually);
addBulkBtn.addEventListener('click', addBulkCoordinates);
plotBtn.addEventListener('click', plotPolygon);
newBtn.addEventListener('click', addNewPolygon);
deleteBtn.addEventListener('click', deleteSelectedPolygon);
refreshBtn.addEventListener('click', refreshAll);
exportKmzBtn.addEventListener('click', exportAsKMZ);
exportShpBtn.addEventListener('click', exportAsSHP);
map.on('click', addCoordinateFromMap);

// Auto-generate numeric serial ID
function generateNextId() {
    return nextId.toString();
}

// Update polygon ID to numeric serial number only
polygonIdInput.addEventListener('blur', function() {
    if (!this.value.trim() || isNaN(this.value.trim())) {
        this.value = generateNextId();
    }
});

// Function to add a coordinate from manual input
function addCoordinateManually() {
    const lat = parseFloat(latInput.value);
    const lng = parseFloat(lngInput.value);

    if (isNaN(lat) || isNaN(lng)) {
        alert('Please enter valid latitude and longitude values');
        return;
    }

    if (lat < -90 || lat > 90) {
        alert('Latitude must be between -90 and 90');
        return;
    }

    if (lng < -180 || lng > 180) {
        alert('Longitude must be between -180 and 180');
        return;
    }

    addCoordinate(lat, lng);

    // Clear input fields
    latInput.value = '';
    lngInput.value = '';
}

// Function to add multiple coordinates at once
function addBulkCoordinates() {
    const coordsText = bulkCoordsInput.value.trim();
    if (!coordsText) {
        alert('Please enter coordinates in the format: lat,lng');
        return;
    }

    // Split by new lines or commas
    const coordLines = coordsText.split(/[\n,]+/).filter(line => line.trim());

    for (const line of coordLines) {
        const parts = line.split(',').map(part => part.trim());
        if (parts.length !== 2) {
            alert(`Invalid coordinate format: ${line}. Please use lat,lng format.`);
            continue;
        }

        const lat = parseFloat(parts[0]);
        const lng = parseFloat(parts[1]);

        if (isNaN(lat) || isNaN(lng)) {
            alert(`Invalid coordinate values: ${line}. Please enter valid numbers.`);
            continue;
        }

        if (lat < -90 || lat > 90) {
            alert(`Latitude out of range in: ${line}. Must be between -90 and 90.`);
            continue;
        }

        if (lng < -180 || lng > 180) {
            alert(`Longitude out of range in: ${line}. Must be between -180 and 180.`);
            continue;
        }

        addCoordinate(lat, lng);
    }

    // Clear the bulk input field
    bulkCoordsInput.value = '';
}

// Function to add a coordinate from map click
function addCoordinateFromMap(e) {
    const lat = e.latlng.lat;
    const lng = e.latlng.lng;
    addCoordinate(lat, lng);
}

// Main function to add a coordinate
function addCoordinate(lat, lng) {
    if (!currentPolygon) {
        // Initialize current polygon if not exists
        currentPolygon = {
            id: polygonIdInput.value.trim() || generateNextId(),
            name: polygonNameInput.value.trim() || `polygon_${nextId}`,
            category: polygonCategorySelect.value,
            points: []
        };
    }

    const latFixed = lat.toFixed(6);
    const lngFixed = lng.toFixed(6);
    currentPoints.push([lat, lng]);

    // Update coordinates display
    updateCoordinatesDisplay();

    // Add marker to map
    const marker = L.marker([lat, lng]).addTo(map)
        .bindPopup(`Point ${currentPoints.length}<br>Lat: ${latFixed}<br>Lng: ${lngFixed}`)
        .openPopup();

    markers.push(marker);

    // Draw polyline if we have at least 2 points
    if (currentPoints.length >= 2) {
        if (currentPolyline) {
            map.removeLayer(currentPolyline);
        }

        // Show a preview polygon when we have at least 3 points
        if (currentPoints.length >= 3) {
            currentPolyline = L.polygon([...currentPoints, currentPoints[0]], {
                color: 'blue',
                weight: 2,
                dashArray: '5, 10',
                fillOpacity: 0.1
            }).addTo(map);
        } else {
            currentPolyline = L.polyline(currentPoints, {
                color: 'blue',
                weight: 2,
                dashArray: '5, 10'
            }).addTo(map);
        }
    }

    // Calculate and update area if we have at least 3 points
    if (currentPoints.length >= 3) {
        updateAreaDisplay();
    }
}

// Update the coordinates display
function updateCoordinatesDisplay() {
    coordinatesDisplay.innerHTML = currentPoints.length > 0 
        ? currentPoints.map((p, i) => `${i+1}. Lat: ${p[0].toFixed(6)}, Lng: ${p[1].toFixed(6)}`).join('<br>')
        : 'No coordinates added yet. Add coordinates or click on the map.';
}

// Calculate and update area display
function updateAreaDisplay() {
    if (currentPoints.length < 3) {
        areaDisplay.textContent = 'Area: 0 m² (0 hectares) (0 sq km)';
        return;
    }

    // Convert points to LatLng objects
    const latLngs = currentPoints.map(p => L.latLng(p[0], p[1]));

    // Calculate area in square meters
    const area = calculateGeodesicArea(latLngs);
    const hectares = area / 10000;
    const sqKm = area / 1000000;

    areaDisplay.textContent = `Area: ${area.toFixed(2)} m² (${hectares.toFixed(4)} hectares) (${sqKm.toFixed(6)} sq km)`;
}

// Plot the current polygon
function plotPolygon() {
    if (currentPoints.length < 3) {
        alert('A polygon requires at least 3 points');
        return;
    }

    // Save state before making changes (for undo/redo)
    if (!isEditMode) {
        saveState();
    }

    // Remove the preview polyline
    if (currentPolyline) {
        map.removeLayer(currentPolyline);
        currentPolyline = null;
    }

    // Ensure ID is numeric only and convert to string for consistent comparison
    let id = polygonIdInput.value.trim();
    if (!id || isNaN(id)) {
        id = generateNextId();
    }
    id = parseInt(id).toString(); // Ensure consistent string format
    
    const name = polygonNameInput.value.trim() || `polygon_${id}`;
    const category = polygonCategorySelect.value;

    // Check if ID already exists (but allow if we're editing the same polygon)
    if (!isEditMode && polygons.find(p => p.id === id)) {
        alert('Polygon ID already exists. Please use a different ID.');
        return;
    }

    // If editing, remove the old polygon first
    if (isEditMode && editingPolygonId) {
        const editIndex = polygons.findIndex(p => p.id === editingPolygonId);
        if (editIndex !== -1) {
            if (polygons[editIndex].layer && polygons[editIndex].layer._map) {
                map.removeLayer(polygons[editIndex].layer);
            }
            polygons.splice(editIndex, 1);
        }
    }

    // Calculate area
    const latLngs = currentPoints.map(p => L.latLng(p[0], p[1]));
    const area = calculateGeodesicArea(latLngs);
    const hectares = area / 10000;
    const sqKm = area / 1000000;

    // Create polygon (closing the loop by adding the first point at the end)
    const polygon = L.polygon([...currentPoints, currentPoints[0]], {
        color: getColorForCategory(category),
        fillOpacity: 0.5,
        weight: 2
    }).addTo(map);

    // Add popup with info
    polygon.bindPopup(`
        <strong>${name}</strong><br>
        ID: ${id}<br>
        Category: ${category}<br>
        Points: ${currentPoints.length}<br>
        Area: ${area.toFixed(2)} m²<br>
        Hectares: ${hectares.toFixed(4)} ha<br>
        Sq Km: ${sqKm.toFixed(6)} sq km
    `);

    // Store polygon data
    const polygonData = {
        id,
        name,
        category,
        points: currentPoints.map(p => [p[0].toFixed(6), p[1].toFixed(6)]),
        latLngs: currentPoints,
        layer: polygon,
        area,
        hectares,
        sqKm,
        attributes: {} // For storing joined Excel data
    };

    polygons.push(polygonData);
    
    if (!isEditMode) {
        nextId++;
    }

    // Update shapes list
    updateShapesList();

    // Reset for new polygon
    isEditMode = false;
    editingPolygonId = null;
    addNewPolygon();
}

// Get color based on category
function getColorForCategory(category) {
    const colors = {
        residential: '#3498db',
        commercial: '#e74c3c',
        industrial: '#2ecc71',
        agricultural: '#f39c12',
        recreational: '#9b59b6'
    };

    return colors[category] || '#3498db';
}

// Add a new polygon (reset form)
function addNewPolygon() {
    // Clear current drawing
    if (currentPolyline) {
        map.removeLayer(currentPolyline);
        currentPolyline = null;
    }

    // Clear markers
    markers.forEach(marker => map.removeLayer(marker));
    markers = [];

    // Clear current points
    currentPoints = [];
    updateCoordinatesDisplay();
    updateAreaDisplay();

    // Generate next numeric ID
    polygonIdInput.value = nextId.toString();
    polygonNameInput.value = `polygon_${nextId}`;
    polygonCategorySelect.value = 'residential';

    // Remove current class from all shapes
    document.querySelectorAll('.shape-item').forEach(item => {
        item.classList.remove('current-shape');
    });

    selectedPolygonId = null;
    currentPolygon = null;
}

// Add polygon to the shapes list
function addToShapesList(polygonData) {
    const shapeItem = document.createElement('div');
    shapeItem.className = 'shape-item';
    shapeItem.dataset.id = polygonData.id;

    shapeItem.innerHTML = `
        <div class="shape-info">
            <strong>${polygonData.name}</strong> (ID: ${polygonData.id})<br>
            <small>${polygonData.category}, ${polygonData.points.length} points</small><br>
            <small>Area: ${polygonData.area.toFixed(2)} m²</small><br>
            <small>Hectares: ${polygonData.hectares.toFixed(4)} ha</small><br>
            <small>Sq Km: ${polygonData.sqKm.toFixed(6)} sq km</small>
        </div>
        <div class="shape-actions">
            <button class="primary" onclick="selectShape('${polygonData.id}')">Select</button>
            <button class="success" onclick="editShape('${polygonData.id}')">Edit</button>
            <button class="danger" onclick="deleteShape('${polygonData.id}')">Delete</button>
        </div>
    `;

    shapesContainer.appendChild(shapeItem);
}

// Select a shape for editing
function selectShape(id) {
    const polygon = polygons.find(p => p.id === id);
    if (!polygon) return;

    selectedPolygonId = id;

    // Remove current class from all shapes
    document.querySelectorAll('.shape-item').forEach(item => {
        item.classList.remove('current-shape');
    });

    // Add current class to selected shape
    const selectedItem = document.querySelector(`.shape-item[data-id="${id}"]`);
    if (selectedItem) {
        selectedItem.classList.add('current-shape');
    }

    // Zoom to the polygon
    map.fitBounds(polygon.layer.getBounds());

    // Highlight the polygon
    polygon.layer.setStyle({
        weight: 4,
        color: '#ffeb3b'
    });

    // Reset style after 3 seconds
    setTimeout(() => {
        if (polygon.layer && polygon.layer._map) {
            polygon.layer.setStyle({
                weight: 2,
                color: getColorForCategory(polygon.category)
            });
        }
    }, 3000);
}

// Edit a specific shape
function editShape(id) {
    const polygon = polygons.find(p => p.id === id);
    if (!polygon) return;

    // Save state for undo
    saveState();

    // Enter edit mode
    isEditMode = true;
    editingPolygonId = id;

    // Clear current drawing
    if (currentPolyline) {
        map.removeLayer(currentPolyline);
        currentPolyline = null;
    }
    markers.forEach(marker => map.removeLayer(marker));
    markers = [];

    // Load polygon data into form
    polygonIdInput.value = polygon.id;
    polygonNameInput.value = polygon.name;
    polygonCategorySelect.value = polygon.category;

    // Load points
    currentPoints = [...polygon.latLngs];
    updateCoordinatesDisplay();
    updateAreaDisplay();

    // Add markers for each point
    currentPoints.forEach((point, index) => {
        const marker = L.marker(point).addTo(map)
            .bindPopup(`Point ${index + 1}<br>Lat: ${point[0].toFixed(6)}<br>Lng: ${point[1].toFixed(6)}`);
        markers.push(marker);
    });

    // Show preview polyline
    if (currentPoints.length >= 3) {
        currentPolyline = L.polygon([...currentPoints, currentPoints[0]], {
            color: 'orange',
            weight: 3,
            dashArray: '5, 10',
            fillOpacity: 0.2
        }).addTo(map);
    }

    // Zoom to the polygon
    map.fitBounds(polygon.layer.getBounds());

    alert('Edit mode activated! Modify coordinates and click "Plot Polygon" to save changes.');
}

// Delete a specific shape
function deleteShape(id) {
    if (!confirm('Are you sure you want to delete this polygon?')) return;

    // Save state for undo
    saveState();

    const index = polygons.findIndex(p => p.id === id);
    if (index === -1) {
        alert('Polygon not found');
        return;
    }

    // Remove from map
    if (polygons[index].layer && polygons[index].layer._map) {
        map.removeLayer(polygons[index].layer);
    }

    // Remove from array
    polygons.splice(index, 1);

    // Clear selection if this was the selected polygon
    if (selectedPolygonId === id) {
        selectedPolygonId = null;
    }

    // Update the shapes list display
    updateShapesList();
    
    alert('Polygon deleted successfully! Use "Undo" if this was a mistake.');
}

// Delete the selected polygon
function deleteSelectedPolygon() {
    if (!selectedPolygonId) {
        alert('Please select a polygon first by clicking the "Select" button next to any polygon in the list');
        return;
    }

    const polygon = polygons.find(p => p.id === selectedPolygonId);
    if (!polygon) {
        alert('Selected polygon not found');
        selectedPolygonId = null;
        return;
    }

    deleteShape(selectedPolygonId);
}

// Refresh all - clear everything
function refreshAll() {
    if (!confirm('Are you sure you want to clear all polygons? This cannot be undone.')) return;

    // Clear all polygons from map
    polygons.forEach(p => map.removeLayer(p.layer));

    // Clear arrays and lists
    polygons = [];
    shapesContainer.innerHTML = '';
    nextId = 1;
    selectedPolygonId = null;
    excelData = null;
    excelColumns = [];

    // Clear Excel status
    document.getElementById('excel-status').textContent = 'No Excel file uploaded';
    document.getElementById('id-column-selector').innerHTML = '<option value="">Select ID Column</option>';

    // Reset current polygon
    addNewPolygon();
}

// Handle Excel file upload
function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            excelData = XLSX.utils.sheet_to_json(worksheet);
            
            // Get column names
            excelColumns = Object.keys(excelData[0] || {});
            updateColumnSelector();
            
            document.getElementById('excel-status').textContent = `Excel file loaded: ${excelData.length} rows`;
        } catch (error) {
            alert('Error reading Excel file: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Update the column selector dropdown
function updateColumnSelector() {
    const selector = document.getElementById('id-column-selector');
    selector.innerHTML = '<option value="">Select ID Column</option>';
    
    excelColumns.forEach(column => {
        const option = document.createElement('option');
        option.value = column;
        option.textContent = column;
        selector.appendChild(option);
    });
}

// Join Excel data with polygons
function joinExcelData() {
    const idColumn = document.getElementById('id-column-selector').value;
    if (!idColumn || !excelData) {
        alert('Please select an ID column and upload an Excel file first');
        return;
    }

    let joinedCount = 0;
    
    polygons.forEach(polygon => {
        const matchingRow = excelData.find(row => row[idColumn]?.toString() === polygon.id);
        if (matchingRow) {
            polygon.attributes = matchingRow;
            joinedCount++;
        }
    });

    alert(`Successfully joined ${joinedCount} polygons with Excel data`);
    updateShapesList();
}

// Update shapes list to show joined data
function updateShapesList() {
    shapesContainer.innerHTML = '';
    polygons.forEach(polygonData => {
        addToShapesList(polygonData);
    });
}

// Export as KMZ with Excel attributes
function exportAsKMZ() {
    if (polygons.length === 0) {
        alert('No polygons to export');
        return;
    }

    let kmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>`;

    polygons.forEach(polygon => {
        let description = `ID: ${polygon.id}, Category: ${polygon.category}, Area: ${polygon.area.toFixed(2)} m², Hectares: ${polygon.hectares.toFixed(4)} ha, Sq Km: ${polygon.sqKm.toFixed(6)} sq km`;
        
        // Add Excel attributes if available
        if (polygon.attributes && Object.keys(polygon.attributes).length > 0) {
            description += '<br><br>Additional Attributes:<br>';
            Object.entries(polygon.attributes).forEach(([key, value]) => {
                description += `${key}: ${value}<br>`;
            });
        }

        kmlContent += `
<Placemark>
<name>${polygon.name}</name>
<description>${description}</description>
<Polygon>
<outerBoundaryIs>
<LinearRing>
<coordinates>
${polygon.points.map(p => `${p[1]},${p[0]},0`).join(' ')}
${polygon.points[0][1]},${polygon.points[0][0]},0
</coordinates>
</LinearRing>
</outerBoundaryIs>
</Polygon>
</Placemark>`;
    });

    kmlContent += `
</Document>
</kml>`;

    // Create blob and download
    const blob = new Blob([kmlContent], {type: 'application/vnd.google-earth.kml+xml'});
    saveAs(blob, 'polygons.kmz');
}

// Export as Shapefile with Excel attributes
function exportAsSHP() {
    if (polygons.length === 0) {
        alert('No polygons to export');
        return;
    }

    try {
        const geojson = {
            type: "FeatureCollection",
            features: polygons.map(polygon => {
                const properties = {
                    id: parseInt(polygon.id),
                    name: polygon.name,
                    category: polygon.category,
                    area_m2: parseFloat(polygon.area.toFixed(2)),
                    area_ha: parseFloat(polygon.hectares.toFixed(4)),
                    area_sqkm: parseFloat(polygon.sqKm.toFixed(6)),
                    points: polygon.points.length
                };

                // Add Excel attributes if available
                if (polygon.attributes && Object.keys(polygon.attributes).length > 0) {
                    Object.assign(properties, polygon.attributes);
                }

                return {
                    type: "Feature",
                    properties,
                    geometry: {
                        type: "Polygon",
                        coordinates: [polygon.points.map(p => [parseFloat(p[1]), parseFloat(p[0])])]
                    }
                };
            })
        };

        const zip = new JSZip();

        // Add GeoJSON file
        zip.file("polygons.geojson", JSON.stringify(geojson, null, 2));
        
        // Add a CSV with attribute data
        let csvContent = "id,name,category,area_m2,area_hectares,area_sqkm,points";
        
        // Add Excel column headers if available
        if (excelColumns.length > 0) {
            csvContent += "," + excelColumns.join(",");
        }
        csvContent += "\n";

        polygons.forEach(polygon => {
            let row = `${polygon.id},"${polygon.name}","${polygon.category}",${polygon.area.toFixed(2)},${polygon.hectares.toFixed(4)},${polygon.sqKm.toFixed(6)},${polygon.points.length}`;
            
            // Add Excel attribute values if available
            if (polygon.attributes && Object.keys(polygon.attributes).length > 0) {
                excelColumns.forEach(column => {
                    row += `,"${polygon.attributes[column] || ''}"`;
                });
            }
            csvContent += row + "\n";
        });
        
        zip.file("attributes.csv", csvContent);

        // Generate and download zip
        zip.generateAsync({type:"blob"}).then(function(content) {
            saveAs(content, "polygons_shapefile.zip");
        });

    } catch (error) {
        console.error('Error exporting Shapefile:', error);
        alert('Error exporting Shapefile. See console for details.');
    }
}

// Make functions globally available
window.selectShape = selectShape;
window.editShape = editShape;
window.deleteShape = deleteShape;
window.handleExcelUpload = handleExcelUpload;
window.joinExcelData = joinExcelData;
window.undo = undo;
window.redo = redo;

// Initialize the application
addNewPolygon();
