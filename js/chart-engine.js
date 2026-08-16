/**
 * 0cell Chart Engine
 * Floating, interactive charts (Column, Bar, Line, Area, Pie, Donut, Scatter)
 * with live data synchronization to spreadsheet cells.
 */

class ChartEngine {
  constructor(app) {
    this.app = app;
    this.chartCounter = 1;
  }

  // Insert chart from active selection
  insertChart(type = 'column') {
    const sheet = this.app.workbook.getActiveSheet();
    const sel = sheet.selection;
    if (!sel) {
      this.app.showToast('Please select a data range for the chart', 'warning');
      return;
    }

    const chartId = `chart_${Date.now()}`;
    const chartData = this.extractChartData(sheet, sel);

    if (chartData.labels.length === 0 || chartData.series.length === 0) {
      this.app.showToast('Selected range does not contain valid data for a chart', 'warning');
      return;
    }

    const chartObj = {
      id: chartId,
      title: chartData.title || `${type.toUpperCase()} Chart ${this.chartCounter++}`,
      type,
      range: { ...sel },
      x: 120 + (sheet.charts.length * 30),
      y: 100 + (sheet.charts.length * 30),
      width: 420,
      height: 260
    };

    sheet.charts.push(chartObj);
    this.renderCharts();
    this.app.showToast(`Inserted ${type} chart`, 'success');
  }

  // Extract labels and series from 2D range
  extractChartData(sheet, range) {
    const { startCol, startRow, endCol, endRow } = range;
    const numRows = endRow - startRow + 1;
    const numCols = endCol - startCol + 1;

    let title = '';
    const labels = [];
    const series = [];

    // Check top-left
    const topLeft = sheet.getCell(startCol, startRow);

    // If multi-column table: 1st column is labels, remaining columns are series
    if (numCols > 1) {
      // Header names
      for (let c = startCol + 1; c <= endCol; c++) {
        const hCell = sheet.getCell(c, startRow);
        series.push({
          name: hCell && hCell.value ? String(hCell.value) : `Series ${c - startCol}`,
          values: []
        });
      }

      // Rows
      for (let r = startRow + 1; r <= endRow; r++) {
        const labelCell = sheet.getCell(startCol, r);
        labels.push(labelCell && labelCell.value !== undefined ? String(labelCell.value) : `Row ${r + 1}`);

        for (let c = startCol + 1; c <= endCol; c++) {
          const sIdx = c - startCol - 1;
          const valCell = sheet.getCell(c, r);
          const val = valCell && valCell.value !== undefined && !isNaN(valCell.value) ? Number(valCell.value) : 0;
          series[sIdx].values.push(val);
        }
      }
    } else {
      // Single column of numbers
      series.push({ name: 'Data', values: [] });
      for (let r = startRow; r <= endRow; r++) {
        labels.push(`Row ${r + 1}`);
        const cCell = sheet.getCell(startCol, r);
        const val = cCell && cCell.value !== undefined && !isNaN(cCell.value) ? Number(cCell.value) : 0;
        series[0].values.push(val);
      }
    }

    return { title, labels, series };
  }

  // Render all active charts on sheet
  renderCharts() {
    const container = document.getElementById('charts-layer');
    if (!container) return;
    container.innerHTML = '';

    const sheet = this.app.workbook.getActiveSheet();
    if (!sheet || !sheet.charts) return;

    sheet.charts.forEach(chart => {
      const data = this.extractChartData(sheet, chart.range);
      const card = document.createElement('div');
      card.className = 'chart-card';
      card.id = chart.id;
      card.style.left = `${chart.x}px`;
      card.style.top = `${chart.y}px`;
      card.style.width = `${chart.width}px`;
      card.style.height = `${chart.height}px`;

      card.innerHTML = `
        <div class="chart-header">
          <span class="chart-title">${chart.title}</span>
          <div class="chart-actions">
            <button class="chart-btn" data-action="delete" title="Delete Chart">✕</button>
          </div>
        </div>
        <div class="chart-canvas-wrap">
          <canvas id="canvas-${chart.id}" width="${chart.width - 24}" height="${chart.height - 50}"></canvas>
        </div>
      `;

      container.appendChild(card);

      // Render chart graphic on canvas
      const canvas = document.getElementById(`canvas-${chart.id}`);
      if (canvas) {
        this.drawChartOnCanvas(canvas, chart.type, data);
      }

      // Drag handling
      this.makeDraggable(card, chart);

      card.querySelector('[data-action="delete"]').onclick = (e) => {
        e.stopPropagation();
        sheet.charts = sheet.charts.filter(c => c.id !== chart.id);
        this.renderCharts();
      };
    });
  }

  makeDraggable(element, chartObj) {
    let isDragging = false;
    let startX, startY, initialLeft, initialTop;

    const header = element.querySelector('.chart-header');
    header.onmousedown = (e) => {
      if (e.target.closest('.chart-btn')) return;
      isDragging = true;
      startX = e.clientX;
      startY = e.clientY;
      initialLeft = chartObj.x;
      initialTop = chartObj.y;
      element.classList.add('selected');

      const onMouseMove = (moveEv) => {
        if (!isDragging) return;
        const dx = moveEv.clientX - startX;
        const dy = moveEv.clientY - startY;
        chartObj.x = Math.max(0, initialLeft + dx);
        chartObj.y = Math.max(0, initialTop + dy);
        element.style.left = `${chartObj.x}px`;
        element.style.top = `${chartObj.y}px`;
      };

      const onMouseUp = () => {
        isDragging = false;
        element.classList.remove('selected');
        window.removeEventListener('mousemove', onMouseMove);
        window.removeEventListener('mouseup', onMouseUp);
      };

      window.addEventListener('mousemove', onMouseMove);
      window.addEventListener('mouseup', onMouseUp);
    };
  }

  // Draw chart graphics
  drawChartOnCanvas(canvas, type, data) {
    const ctx = canvas.getContext('2d');
    const w = canvas.width;
    const h = canvas.height;
    ctx.clearRect(0, 0, w, h);

    const colors = ['#107c41', '#0078d4', '#d83b01', '#5c2d91', '#008272', '#ffb900'];

    if (type === 'column') {
      this.drawColumnChart(ctx, w, h, data, colors);
    } else if (type === 'bar') {
      this.drawBarChart(ctx, w, h, data, colors);
    } else if (type === 'line' || type === 'area') {
      this.drawLineChart(ctx, w, h, data, colors, type === 'area');
    } else if (type === 'pie' || type === 'donut') {
      this.drawPieChart(ctx, w, h, data, colors, type === 'donut');
    } else {
      this.drawColumnChart(ctx, w, h, data, colors);
    }
  }

  drawColumnChart(ctx, w, h, data, colors) {
    const padding = { top: 20, right: 20, bottom: 35, left: 45 };
    const chartW = w - padding.left - padding.right;
    const chartH = h - padding.top - padding.bottom;

    // Find max value
    let maxVal = 10;
    data.series.forEach(s => {
      s.values.forEach(v => { if (v > maxVal) maxVal = v; });
    });
    maxVal = Math.ceil(maxVal * 1.15);

    // Draw axes & grid
    ctx.strokeStyle = '#e1dfdd';
    ctx.fillStyle = '#616161';
    ctx.font = '10px Segoe UI';
    ctx.lineWidth = 1;

    for (let i = 0; i <= 4; i++) {
      const y = padding.top + (chartH / 4) * i;
      const val = Math.round(maxVal - (maxVal / 4) * i);
      ctx.beginPath();
      ctx.moveTo(padding.left, y);
      ctx.lineTo(w - padding.right, y);
      ctx.stroke();
      ctx.fillText(val, 6, y + 3);
    }

    // Draw bars
    const numGroups = data.labels.length;
    const numSeries = data.series.length;
    const groupWidth = chartW / numGroups;
    const barWidth = Math.max(4, (groupWidth * 0.7) / numSeries);

    data.labels.forEach((lbl, gIdx) => {
      const groupX = padding.left + gIdx * groupWidth;

      // Label
      ctx.fillStyle = '#616161';
      ctx.textAlign = 'center';
      ctx.fillText(lbl.substring(0, 10), groupX + groupWidth / 2, h - 12);

      data.series.forEach((ser, sIdx) => {
        const val = ser.values[gIdx] || 0;
        const barH = (val / maxVal) * chartH;
        const barX = groupX + (groupWidth * 0.15) + sIdx * barWidth;
        const barY = padding.top + chartH - barH;

        ctx.fillStyle = colors[sIdx % colors.length];
        ctx.fillRect(barX, barY, barWidth - 2, barH);
      });
    });
  }

  drawBarChart(ctx, w, h, data, colors) {
    const padding = { top: 20, right: 30, bottom: 25, left: 60 };
    const chartW = w - padding.left - padding.right;
    const chartH = h - padding.top - padding.bottom;

    let maxVal = 10;
    data.series.forEach(s => s.values.forEach(v => { if (v > maxVal) maxVal = v; }));
    maxVal = Math.ceil(maxVal * 1.15);

    const numRows = data.labels.length;
    const rowH = chartH / numRows;

    data.labels.forEach((lbl, rIdx) => {
      const y = padding.top + rIdx * rowH;
      ctx.fillStyle = '#616161';
      ctx.textAlign = 'right';
      ctx.font = '10px Segoe UI';
      ctx.fillText(lbl.substring(0, 8), padding.left - 6, y + rowH / 2 + 3);

      const val = data.series[0].values[rIdx] || 0;
      const barW = (val / maxVal) * chartW;

      ctx.fillStyle = colors[0];
      ctx.fillRect(padding.left, y + rowH * 0.15, barW, rowH * 0.7);

      ctx.textAlign = 'left';
      ctx.fillText(val, padding.left + barW + 4, y + rowH / 2 + 3);
    });
  }

  drawLineChart(ctx, w, h, data, colors, isArea = false) {
    const padding = { top: 20, right: 20, bottom: 35, left: 45 };
    const chartW = w - padding.left - padding.right;
    const chartH = h - padding.top - padding.bottom;

    let maxVal = 10;
    data.series.forEach(s => s.values.forEach(v => { if (v > maxVal) maxVal = v; }));
    maxVal = Math.ceil(maxVal * 1.15);

    // Axes
    ctx.strokeStyle = '#e1dfdd';
    ctx.lineWidth = 1;
    for (let i = 0; i <= 4; i++) {
      const y = padding.top + (chartH / 4) * i;
      const val = Math.round(maxVal - (maxVal / 4) * i);
      ctx.beginPath();
      ctx.moveTo(padding.left, y);
      ctx.lineTo(w - padding.right, y);
      ctx.stroke();
      ctx.fillStyle = '#616161';
      ctx.textAlign = 'left';
      ctx.fillText(val, 6, y + 3);
    }

    const numPts = data.labels.length;
    const stepX = chartW / Math.max(1, numPts - 1);

    data.series.forEach((ser, sIdx) => {
      const color = colors[sIdx % colors.length];
      ctx.strokeStyle = color;
      ctx.fillStyle = color;
      ctx.lineWidth = 2;

      ctx.beginPath();
      ser.values.forEach((v, idx) => {
        const px = padding.left + idx * stepX;
        const py = padding.top + chartH - (v / maxVal) * chartH;
        if (idx === 0) ctx.moveTo(px, py);
        else ctx.lineTo(px, py);
      });
      ctx.stroke();

      if (isArea) {
        ctx.lineTo(padding.left + (ser.values.length - 1) * stepX, padding.top + chartH);
        ctx.lineTo(padding.left, padding.top + chartH);
        ctx.closePath();
        ctx.fillStyle = color + '22';
        ctx.fill();
      }

      // Draw points
      ser.values.forEach((v, idx) => {
        const px = padding.left + idx * stepX;
        const py = padding.top + chartH - (v / maxVal) * chartH;
        ctx.beginPath();
        ctx.arc(px, py, 3, 0, Math.PI * 2);
        ctx.fillStyle = color;
        ctx.fill();
      });
    });

    // Labels
    data.labels.forEach((lbl, idx) => {
      const px = padding.left + idx * stepX;
      ctx.fillStyle = '#616161';
      ctx.textAlign = 'center';
      ctx.fillText(lbl.substring(0, 8), px, h - 12);
    });
  }

  drawPieChart(ctx, w, h, data, colors, isDonut = false) {
    const cx = w / 2;
    const cy = h / 2 + 5;
    const radius = Math.min(cx, cy) - 30;

    const values = data.series[0] ? data.series[0].values : [];
    const total = values.reduce((a, b) => a + b, 0) || 1;

    let startAngle = -Math.PI / 2;

    values.forEach((v, idx) => {
      const sliceAngle = (v / total) * (Math.PI * 2);
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, radius, startAngle, startAngle + sliceAngle);
      ctx.closePath();
      ctx.fillStyle = colors[idx % colors.length];
      ctx.fill();
      startAngle += sliceAngle;
    });

    if (isDonut) {
      ctx.beginPath();
      ctx.arc(cx, cy, radius * 0.55, 0, Math.PI * 2);
      ctx.fillStyle = '#ffffff';
      ctx.fill();
    }
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = ChartEngine;
}
