import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import * as XLSX from 'xlsx';

interface ExcelRow {
  [key: string]: any;
  rowNumber: number;
  // Columnas calculadas
  '% UTIL GANADA'?: number;
  'MONTO GANADO'?: number;
  'COLUMNA 3'?: number;
}

@Component({
  selector: 'app-agent-commission',
  imports: [CommonModule],
  templateUrl: './agent-commission.component.html',
  styleUrl: './agent-commission.component.scss'
})
export class AgentCommissionComponent {

  selectedFile: File | null = null;
  excelData: ExcelRow[] = [];
  allHeaders: string[] = [];
  displayedHeaders: string[] = [];
  workbook: any = null;
  sheetName: string = '';

  // Columnas que queremos mostrar
  private readonly columnsToShow = [
    'CodArt',
    'Articulo',
    'Cantidad#sumar',
    'TOTAL sin IVA#sumar',
    'UTIL_porc',
    'Vendedor',
    '% UTIL GANADA',
    'MONTO GANADO',
    'COLUMNA 3'
  ];

  onFileSelected(event: any): void {
    const file: File = event.target.files[0];
    if (file && this.isExcelFile(file)) {
      this.selectedFile = file;
    } else {
      alert('Por favor, selecciona un archivo Excel válido (.xlsx, .xls)');
      this.selectedFile = null;
    }
  }

  readExcel(): void {
    if (!this.selectedFile) return;

    const reader = new FileReader();
    reader.onload = (e: any) => {
      try {
        const data = new Uint8Array(e.target.result);
        this.workbook = XLSX.read(data, { type: 'array' });

        // Leer la primera hoja
        this.sheetName = this.workbook.SheetNames[0];
        const worksheet = this.workbook.Sheets[this.sheetName];

        // Convertir a JSON con objetos (no arrays)
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        // Procesar los datos para una mejor estructura
        this.processExcelData(jsonData);

      } catch (error) {
        console.error('Error al leer el archivo:', error);
        alert('Error al leer el archivo. Asegúrate de que es un Excel válido.');
      }
    };

    reader.onerror = (error) => {
      console.error('Error al leer el archivo:', error);
      alert('Error al leer el archivo.');
    };

    reader.readAsArrayBuffer(this.selectedFile);
  }

  private processExcelData(rawData: any[]): void {
    this.excelData = [];
    this.allHeaders = [];

    if (rawData.length === 0) return;

    // Obtener headers limpios
    const originalHeaders = Object.keys(rawData[0]);
    this.allHeaders = originalHeaders.map(header => header.trim());

    // Procesar cada fila
    rawData.forEach((row, index) => {
      const processedRow: ExcelRow = { rowNumber: index + 1 };

      // Usar el índice para mapear correctamente
      originalHeaders.forEach((originalHeader, i) => {
        const cleanHeader = this.allHeaders[i];
        processedRow[cleanHeader] = row[originalHeader];
      });

      // Calcular las columnas adicionales
      this.calculateAdditionalColumns(processedRow);

      this.excelData.push(processedRow);
    });

    // Filtrar solo las columnas que queremos mostrar
    this.displayedHeaders = this.columnsToShow.filter(header =>
      this.allHeaders.includes(header) ||
      ['% UTIL GANADA', 'MONTO GANADO', 'COLUMNA 3'].includes(header)
    );
  }

  private calculateAdditionalColumns(row: ExcelRow): void {
    // Obtener los valores necesarios para los cálculos
    const utilPorc = this.getNumberValue(row['UTIL_porc']);
    const totalSinIva = this.getNumberValue(row['TOTAL sin IVA#sumar']);

    // 1. Calcular % UTIL GANADA
    row['% UTIL GANADA'] = this.calculateUtilGanada(utilPorc);

    // 2. Calcular MONTO GANADO
    row['MONTO GANADO'] = totalSinIva * (row['% UTIL GANADA'] || 0);

    // 3. Calcular COLUMNA TRES
    row['COLUMNA 3'] = (row['MONTO GANADO'] || 0) / 1.5;
  }

  private calculateUtilGanada(utilPorc: number): number {
    // =+SI([@[% Utilidad]]<=5,0,SI([@[% Utilidad]]<=9,0.0015,SI([@[% Utilidad]]<=19,0.007,SI([@[% Utilidad]]<=38,0.015,SI([@[% Utilidad]]<=63,0.03,SI([@[% Utilidad]]<=99,0.05))))))
    if (utilPorc <= 5) return 0;
    if (utilPorc <= 9) return 0.0015;
    if (utilPorc <= 19) return 0.007;
    if (utilPorc <= 38) return 0.015;
    if (utilPorc <= 63) return 0.03;
    if (utilPorc <= 99) return 0.05;
    return 0.05; // Para valores mayores a 99
  }

  private getNumberValue(value: any): number {
    if (value === null || value === undefined || value === '') return 0;
    const num = Number(value);
    return isNaN(num) ? 0 : num;
  }

  formatCellValue(value: any): string {
    if (value === null || value === undefined || value === '') {
      return '-';
    }

    // Formatear números
    if (typeof value === 'number') {
      // Para porcentajes
      if (value < 1 && value > 0) {
        return (value * 100).toFixed(2) + '%';
      }

      // Para montos de dinero
      if (Math.abs(value) >= 1) {
        return value.toLocaleString('es-ES', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2
        });
      }

      // Números pequeños
      return value.toFixed(4);
    }

    // Formatear fechas
    if (value instanceof Date) {
      return value.toLocaleDateString('es-ES');
    }

    // Para strings, trim y mostrar
    return String(value).trim();
  }

  isNumber(value: any): boolean {
    return typeof value === 'number';
  }

  getTotalMontoGanado(): number {
    return this.excelData.reduce((total, row) => total + (row['MONTO GANADO'] || 0), 0);
  }

  getSheetName(): string {
    return this.sheetName || 'N/A';
  }

  deleteRow(index: number): void {
    if (confirm('¿Estás seguro de que quieres eliminar esta fila?')) {
      this.excelData.splice(index, 1);
      // Recalcular números de fila
      this.excelData.forEach((row, i) => row.rowNumber = i + 1);
    }
  }

  clearData(): void {
    this.excelData = [];
    this.allHeaders = [];
    this.displayedHeaders = [];
    this.selectedFile = null;
    this.workbook = null;
    this.sheetName = '';
    const fileInput = document.getElementById('excelFile') as HTMLInputElement;
    if (fileInput) {
      fileInput.value = '';
    }
  }

  private isExcelFile(file: File): boolean {
    return file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
  }

  getTotalMontoGanadoSum(): number {
    return this.excelData.reduce((total, row) => total + (row['MONTO GANADO'] || 0), 0);
  }

  getTotalColumnaTresSum(): number {
    return this.excelData.reduce((total, row) => total + (row['COLUMNA 3'] || 0), 0);
  }

  exportToExcelWithAdvancedFormatting(): void {
    const agent = this.getAgentName();
    const month = this.getMonth();
    const year = new Date().getFullYear();
    const fileName = `${agent} - ${month} ${year}`;


    // Usar array de arrays para mejor control
    const data: any[][] = [];

    // 1. Encabezados
    data.push(this.displayedHeaders);

    // 2. Datos
    this.excelData.forEach(row => {
      const rowData: any[] = [];
      this.displayedHeaders.forEach(header => {
        rowData.push(row[header]);
      });
      data.push(rowData);
    });

    // 3. Fila de totales
    const totalRow: any[] = [];
    this.displayedHeaders.forEach(header => {
      if (header === 'MONTO GANADO') {
        totalRow.push(this.getTotalMontoGanado());
      } else if (header === 'COLUMNA TRES') {
        totalRow.push(this.getTotalColumnaTresSum());
      } else if (header === 'CodArt') {
        totalRow.push('TOTAL');
      } else {
        totalRow.push('');
      }
    });
    data.push(totalRow);

    // Crear hoja
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(data);

    // Aplicar formatos
    this.applyAdvancedExcelFormatting(ws, data.length);

    // Exportar con nombre personalizado 🔥
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Reporte');
    XLSX.writeFile(wb, `${fileName}.xlsx`);
  }

  private getAgentName(): string {
    if (this.excelData.length === 0) return 'SinVendedor';

    const firstAgent = this.excelData[0]['Vendedor'];

    return firstAgent ? firstAgent.toString().trim() : 'SinVendedor';
  }

  private getMonth(): string {
    const months = [
      'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
      'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
    ];

    const nowDate = new Date();
    return months[nowDate.getMonth()];
  }


  private applyAdvancedExcelFormatting(ws: XLSX.WorkSheet, totalRows: number): void {
    const lastRow = totalRows - 1; // Índice base 0

    // Formato para encabezados (fila 0)
    this.displayedHeaders.forEach((_, colIndex) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: colIndex });
      ws[cellAddress].s = {
        font: { bold: true, color: { rgb: "FFFFFFFF" } },
        fill: { fgColor: { rgb: "FF0070C0" } }, // Azul
        alignment: { horizontal: 'center' }
      };
    });

    // Formato para totales (última fila)
    this.displayedHeaders.forEach((header, colIndex) => {
      const cellAddress = XLSX.utils.encode_cell({ r: lastRow, c: colIndex });

      if (ws[cellAddress]) {
        ws[cellAddress].s = {
          font: { bold: true },
          fill: { fgColor: { rgb: "FFF2F2F2" } },
          border: {
            top: { style: 'medium', color: { rgb: "FF000000" } }
          }
        };

        // Formatos numéricos
        if (header === 'MONTO GANADO' || header === 'COLUMNA TRES') {
          ws[cellAddress].s.numFmt = '"$"#,##0.00';
        } else if (header.includes('%') || header === 'UTIL_porc') {
          ws[cellAddress].s.numFmt = '0.00%';
        }
      }
    });

    // Ajustar anchos
    ws['!cols'] = this.displayedHeaders.map(header => ({
      wch: Math.max(header.length + 4,
        header === 'Articulo' ? 40 :
          header === 'Vendedor' ? 30 : 15)
    }));
  }


}
