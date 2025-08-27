
// Initialize the map
const map = L.map('map').setView([0, 0], 2);

// Add OpenStreetMap tile layer
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
}).addTo(map);

// Global variables
let excelData = [];
let excelColumns = [];
let pointConfigs = [];
let plottedAreas = [];
let plotType = 'area';
let undoStack = [];
let redoStack = [];

// Geodesic area calculation function
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

// Calculate distance between two points
function calculateDistance(lat1, lng1, lat2, lng2) {
    const R = 6371000; // Earth's radius in meters
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLng = (lng2 - lng1) * Math.PI / 180;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
              Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
              Math.sin(dLng/2) * Math.sin(dLng/2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
    return R * c;
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
            
            document.getElementById('file-status').innerHTML = `
                <strong>File loaded successfully!</strong><br>
                Rows: ${excelData.length}<br>
                Columns: ${excelColumns.length}<br>
                Available columns: ${excelColumns.join(', ')}
            `;
            
            // Enable configuration panels
            document.getElementById('plot-config').classList.remove('disabled');
            document.getElementById('points-config').classList.remove('disabled');
            document.getElementById('plot-controls').classList.remove('disabled');
            
            // Initialize with first point configuration
            initializePoints();
            
        } catch (error) {
            alert('Error reading Excel file: ' + error.message);
            document.getElementById('file-status').textContent = 'Error loading file';
        }
    };
    reader.readAsArrayBuffer(file);
}

// Update plot type
function updatePlotType() {
    plotType = document.querySelector('input[name="plotType"]:checked').value;
    updatePointsInterface();
}

// Initialize points configuration
function initializePoints() {
    pointConfigs = [];
    addPoint();
    addPoint();
}

// Add a new point configuration
function addPoint() {
    const pointNumber = pointConfigs.length + 1;
    const pointConfig = {
        id: pointNumber,
        latColumn: '',
        lngColumn: ''
    };
    
    pointConfigs.push(pointConfig);
    updatePointsInterface();
}

// Remove the last point configuration
function removeLastPoint() {
    if (pointConfigs.length > 2) {
        pointConfigs.pop();
        updatePointsInterface();
    } else {
        alert('Minimum 2 points required');
    }
}

// Update the points interface
function updatePointsInterface() {
    const container = document.getElementById('points-container');
    container.innerHTML = '';
    
    pointConfigs.forEach((config, index) => {
        const pointDiv = document.createElement('div');
        pointDiv.className = 'point-config';
        
        pointDiv.innerHTML = `
            <div class="point-header">
                <h3>Point ${config.id}</h3>
            </div>
            <div class="column-selector">
                <div>
                    <label>Latitude Column</label>
                    <select onchange="updatePointConfig(${index}, 'lat', this.value)">
                        <option value="">Select latitude column</option>
                        ${excelColumns.map(col => 
                            `<option value="${col}" ${config.latColumn === col ? 'selected' : ''}>${col}</option>`
                        ).join('')}
                    </select>
                </div>
                <div>
                    <label>Longitude Column</label>
                    <select onchange="updatePointConfig(${index}, 'lng', this.value)">
                        <option value="">Select longitude column</option>
                        ${excelColumns.map(col => 
                            `<option value="${col}" ${config.lngColumn === col ? 'selected' : ''}>${col}</option>`
                        ).join('')}
                    </select>
                </div>
            </div>
        `;
        
        container.appendChild(pointDiv);
    });
}

// Update point configuration
function updatePointConfig(index, type, value) {
    if (type === 'lat') {
        pointConfigs[index].latColumn = value;
    } else if (type === 'lng') {
        pointConfigs[index].lngColumn = value;
    }
}

// Save state for undo/redo
function saveState() {
    const state = {
        plottedAreas: JSON.parse(JSON.stringify(plottedAreas.map(area => ({
            id: area.id,
            points: area.points,
            area: area.area,
            hectares: area.hectares,
            sqKm: area.sqKm,
            attributes: area.attributes,
            type: area.type
        }))))
    };
    undoStack.push(state);
    if (undoStack.length > 20) {
        undoStack.shift();
    }
    redoStack = [];
}

// Undo function
function undo() {
    if (undoStack.length === 0) {
        alert('Nothing to undo');
        return;
    }
    
    const currentState = {
        plottedAreas: JSON.parse(JSON.stringify(plottedAreas.map(area => ({
            id: area.id,
            points: area.points,
            area: area.area,
            hectares: area.hectares,
            sqKm: area.sqKm,
            attributes: area.attributes,
            type: area.type
        }))))
    };
    redoStack.push(currentState);
    
    const previousState = undoStack.pop();
    restoreState(previousState);
}

// Redo function
function redo() {
    if (redoStack.length === 0) {
        alert('Nothing to redo');
        return;
    }
    
    const currentState = {
        plottedAreas: JSON.parse(JSON.stringify(plottedAreas.map(area => ({
            id: area.id,
            points: area.points,
            area: area.area,
            hectares: area.hectares,
            sqKm: area.sqKm,
            attributes: area.attributes,
            type: area.type
        }))))
    };
    undoStack.push(currentState);
    
    const nextState = redoStack.pop();
    restoreState(nextState);
}

// Restore state
function restoreState(state) {
    // Clear existing areas from map
    plottedAreas.forEach(area => {
        if (area.layer && area.layer._map) {
            map.removeLayer(area.layer);
        }
    });
    
    plottedAreas = [];
    
    // Restore areas
    state.plottedAreas.forEach(areaData => {
        if (areaData.type === 'area' && areaData.points.length >= 3) {
            const latLngs = areaData.points.map(p => L.latLng(p[0], p[1]));
            const polygon = L.polygon(latLngs, {
                color: '#3498db',
                fillOpacity: 0.5,
                weight: 2
            }).addTo(map);
            
            polygon.bindPopup(createPopupContent(areaData));
            areaData.layer = polygon;
            plottedAreas.push(areaData);
        }
    });
    
    updateAreasList();
}

// Plot data from Excel
function plotData() {
    // Validate point configurations
    const isValid = pointConfigs.every(config => config.latColumn && config.lngColumn);
    if (!isValid) {
        alert('Please select latitude and longitude columns for all points');
        return;
    }
    
    saveState();
    
    let plotCount = 0;
    let skippedCount = 0;
    
    excelData.forEach((row, rowIndex) => {
        const points = [];
        
        // Extract coordinates for each configured point
        pointConfigs.forEach(config => {
            const lat = parseFloat(row[config.latColumn]);
            const lng = parseFloat(row[config.lngColumn]);
            
            // Only add valid coordinates (ignore empty cells)
            if (!isNaN(lat) && !isNaN(lng) && lat !== 0 && lng !== 0) {
                points.push([lat, lng]);
            }
        });
        
        // Plot area if we have at least 3 points
        if (plotType === 'area' && points.length >= 3) {
            plotArea(points, row, rowIndex + 1);
            plotCount++;
        } else if (plotType === 'distance' && points.length >= 2) {
            calculateRowDistances(points, row, rowIndex + 1);
            plotCount++;
        } else {
            skippedCount++;
        }
    });
    
    document.getElementById('plot-status').innerHTML = `
        Plotted: ${plotCount} rows<br>
        Skipped: ${skippedCount} rows (insufficient points)<br>
        Total areas: ${plottedAreas.length}
    `;
    
    updateAreasList();
    
    // Zoom to fit all plotted areas
    if (plottedAreas.length > 0) {
        const group = new L.featureGroup(plottedAreas.map(area => area.layer).filter(layer => layer));
        map.fitBounds(group.getBounds().pad(0.1));
    }
}

// Plot individual area
function plotArea(points, attributes, rowId) {
    const latLngs = points.map(p => L.latLng(p[0], p[1]));
    const area = calculateGeodesicArea(latLngs);
    const hectares = area / 10000;
    const sqKm = area / 1000000;
    
    const polygon = L.polygon(latLngs, {
        color: '#3498db',
        fillOpacity: 0.5,
        weight: 2
    }).addTo(map);
    
    const areaData = {
        id: rowId,
        points: points,
        area: area,
        hectares: hectares,
        sqKm: sqKm,
        attributes: attributes,
        layer: polygon,
        type: 'area'
    };
    
    polygon.bindPopup(createPopupContent(areaData));
    plottedAreas.push(areaData);
}

// Calculate distances for a row
function calculateRowDistances(points, attributes, rowId) {
    let totalDistance = 0;
    const distances = [];
    
    for (let i = 0; i < points.length - 1; i++) {
        const dist = calculateDistance(
            points[i][0], points[i][1],
            points[i + 1][0], points[i + 1][1]
        );
        distances.push(dist);
        totalDistance += dist;
    }
    
    // Create a polyline for visualization
    const polyline = L.polyline(points, {
        color: '#e74c3c',
        weight: 3
    }).addTo(map);
    
    const distanceData = {
        id: rowId,
        points: points,
        distances: distances,
        totalDistance: totalDistance,
        attributes: attributes,
        layer: polyline,
        type: 'distance'
    };
    
    polyline.bindPopup(createDistancePopupContent(distanceData));
    plottedAreas.push(distanceData);
}

// Create popup content for areas
function createPopupContent(areaData) {
    let content = `
        <strong>Area ${areaData.id}</strong><br>
        Points: ${areaData.points.length}<br>
        Area: ${areaData.area.toFixed(2)} m²<br>
        Hectares: ${areaData.hectares.toFixed(4)} ha<br>
        Sq Km: ${areaData.sqKm.toFixed(6)} sq km<br><br>
        <strong>Attributes:</strong><br>
    `;
    
    Object.entries(areaData.attributes).forEach(([key, value]) => {
        if (value !== undefined && value !== null && value !== '') {
            content += `${key}: ${value}<br>`;
        }
    });
    
    return content;
}

// Create popup content for distances
function createDistancePopupContent(distanceData) {
    let content = `
        <strong>Distance Analysis ${distanceData.id}</strong><br>
        Points: ${distanceData.points.length}<br>
        Total Distance: ${(distanceData.totalDistance / 1000).toFixed(3)} km<br><br>
        <strong>Segment Distances:</strong><br>
    `;
    
    distanceData.distances.forEach((dist, index) => {
        content += `Segment ${index + 1}: ${dist.toFixed(2)} m<br>`;
    });
    
    content += '<br><strong>Attributes:</strong><br>';
    Object.entries(distanceData.attributes).forEach(([key, value]) => {
        if (value !== undefined && value !== null && value !== '') {
            content += `${key}: ${value}<br>`;
        }
    });
    
    return content;
}

// Update areas list
function updateAreasList() {
    const container = document.getElementById('areas-list');
    container.innerHTML = '';
    
    plottedAreas.forEach(area => {
        const areaDiv = document.createElement('div');
        areaDiv.className = 'area-item';
        
        if (area.type === 'area') {
            areaDiv.innerHTML = `
                <strong>Area ${area.id}</strong><br>
                ${area.points.length} points, ${area.area.toFixed(2)} m²<br>
                ${area.hectares.toFixed(4)} ha, ${area.sqKm.toFixed(6)} sq km
            `;
        } else {
            areaDiv.innerHTML = `
                <strong>Distance ${area.id}</strong><br>
                ${area.points.length} points, ${(area.totalDistance / 1000).toFixed(3)} km total
            `;
        }
        
        container.appendChild(areaDiv);
    });
}

// Clear all plotted data
function clearAll() {
    if (!confirm('Are you sure you want to clear all plotted data?')) return;
    
    plottedAreas.forEach(area => {
        if (area.layer && area.layer._map) {
            map.removeLayer(area.layer);
        }
    });
    
    plottedAreas = [];
    undoStack = [];
    redoStack = [];
    
    updateAreasList();
    document.getElementById('plot-status').textContent = 'All data cleared';
}

// Export as KMZ
function exportKMZ() {
    if (plottedAreas.length === 0) {
        alert('No data to export');
        return;
    }
    
    let kmlContent = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
<name>Resscomm Excel Plotter Results</name>`;

    plottedAreas.forEach(item => {
        if (item.type === 'area') {
            let description = `Area: ${item.area.toFixed(2)} m², Hectares: ${item.hectares.toFixed(4)} ha, Sq Km: ${item.sqKm.toFixed(6)} sq km`;
            
            // Add attributes
            Object.entries(item.attributes).forEach(([key, value]) => {
                if (value !== undefined && value !== null && value !== '') {
                    description += `<br>${key}: ${value}`;
                }
            });

            kmlContent += `
<Placemark>
<name>Area ${item.id}</name>
<description>${description}</description>
<Polygon>
<outerBoundaryIs>
<LinearRing>
<coordinates>
${item.points.map(p => `${p[1]},${p[0]},0`).join(' ')}
${item.points[0][1]},${item.points[0][0]},0
</coordinates>
</LinearRing>
</outerBoundaryIs>
</Polygon>
</Placemark>`;
        } else {
            let description = `Total Distance: ${(item.totalDistance / 1000).toFixed(3)} km`;
            
            // Add attributes
            Object.entries(item.attributes).forEach(([key, value]) => {
                if (value !== undefined && value !== null && value !== '') {
                    description += `<br>${key}: ${value}`;
                }
            });

            kmlContent += `
<Placemark>
<name>Distance ${item.id}</name>
<description>${description}</description>
<LineString>
<coordinates>
${item.points.map(p => `${p[1]},${p[0]},0`).join(' ')}
</coordinates>
</LineString>
</Placemark>`;
        }
    });

    kmlContent += `
</Document>
</kml>`;

    const blob = new Blob([kmlContent], {type: 'application/vnd.google-earth.kml+xml'});
    saveAs(blob, 'resscomm_excel_results.kmz');
}

// Export as Shapefile
function exportShapefile() {
    if (plottedAreas.length === 0) {
        alert('No data to export');
        return;
    }

    const geojson = {
        type: "FeatureCollection",
        features: plottedAreas.map(item => {
            const properties = {
                id: item.id,
                type: item.type
            };

            if (item.type === 'area') {
                properties.area_m2 = parseFloat(item.area.toFixed(2));
                properties.area_ha = parseFloat(item.hectares.toFixed(4));
                properties.area_sqkm = parseFloat(item.sqKm.toFixed(6));
                properties.points = item.points.length;
            } else {
                properties.total_dist_m = parseFloat(item.totalDistance.toFixed(2));
                properties.total_dist_km = parseFloat((item.totalDistance / 1000).toFixed(3));
                properties.points = item.points.length;
            }

            // Add Excel attributes
            Object.assign(properties, item.attributes);

            const geometry = item.type === 'area' ? {
                type: "Polygon",
                coordinates: [item.points.map(p => [parseFloat(p[1]), parseFloat(p[0])])]
            } : {
                type: "LineString",
                coordinates: item.points.map(p => [parseFloat(p[1]), parseFloat(p[0])])
            };

            return {
                type: "Feature",
                properties,
                geometry
            };
        })
    };

    const zip = new JSZip();
    zip.file("resscomm_excel_results.geojson", JSON.stringify(geojson, null, 2));
    
    // Create CSV with all data
    let csvContent = "id,type";
    if (plottedAreas.some(a => a.type === 'area')) {
        csvContent += ",area_m2,area_hectares,area_sqkm";
    }
    if (plottedAreas.some(a => a.type === 'distance')) {
        csvContent += ",total_distance_m,total_distance_km";
    }
    csvContent += ",points";
    
    // Add Excel column headers
    if (excelColumns.length > 0) {
        csvContent += "," + excelColumns.join(",");
    }
    csvContent += "\n";

    plottedAreas.forEach(item => {
        let row = `${item.id},${item.type}`;
        
        if (plottedAreas.some(a => a.type === 'area')) {
            if (item.type === 'area') {
                row += `,${item.area.toFixed(2)},${item.hectares.toFixed(4)},${item.sqKm.toFixed(6)}`;
            } else {
                row += ",,,";
            }
        }
        
        if (plottedAreas.some(a => a.type === 'distance')) {
            if (item.type === 'distance') {
                row += `,${item.totalDistance.toFixed(2)},${(item.totalDistance / 1000).toFixed(3)}`;
            } else {
                row += ",,";
            }
        }
        
        row += `,${item.points.length}`;
        
        // Add Excel attribute values
        excelColumns.forEach(column => {
            row += `,"${item.attributes[column] || ''}"`;
        });
        
        csvContent += row + "\n";
    });
    
    zip.file("attributes.csv", csvContent);

    zip.generateAsync({type:"blob"}).then(function(content) {
        saveAs(content, "resscomm_excel_results_shapefile.zip");
    });
}

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    // Set initial state
    document.getElementById('plot-config').classList.add('disabled');
    document.getElementById('points-config').classList.add('disabled');
    document.getElementById('plot-controls').classList.add('disabled');
});
