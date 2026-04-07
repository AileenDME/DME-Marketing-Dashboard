import { parseISO, format } from 'date-fns';

class ExcelDataService {
  static async fetchFromSharePoint() {
    try {
      // This would call your backend API that handles the Microsoft Graph auth
      // For now, we'll use a CORS proxy or fetch directly if configured
      const response = await fetch('/api/excel-data');
      return await response.json();
    } catch (error) {
      console.error('Error fetching from SharePoint:', error);
      throw error;
    }
  }

  static async loadSampleData() {
    // Sample data for demo purposes
    return {
      leadsGenerated: this.generateSampleData(20, 50),
      leadsVelocity: this.generateSampleDataPercentage(20, 0.5, 0.8),
      domainAuthority: this.generateSampleData(20, 40),
      usTraffic: this.generateSampleData(20, 3000),
      linkedinFollowers: this.generateSampleData(20, 1500),
      salesDeck: this.generateSampleData(20, 3),
      rfpResponses: this.generateSampleData(20, 2),
      projects: [
        { name: 'athenahealth Marketplace', percentage: 0.85, status: 'On Track' },
        { name: 'eClinicalWorks Alliance', percentage: 0.60, status: 'On Track' },
        { name: 'Epic Marketplace Application', percentage: 0.25, status: 'At Risk' },
        { name: 'Q2 Blog Strategy', percentage: 0.40, status: 'On Track' }
      ],
      sharePointLink: 'https://dmeservicesolutions-my.sharepoint.com/:x:/g/personal/aileen_manalo_dmeserve_com/IQBIgn_igqUHQa0xmaavGaToAeuVpToC4cHb8IrqkN-JryQ'
    };
  }

  static generateSampleData(days, maxValue) {
    const data = [];
    const today = new Date();
    for (let i = 0; i < days; i++) {
      const date = new Date(today);
      date.setDate(date.getDate() - days + i);
      data.push({
        date: format(date, 'yyyy-MM-dd'),
        value: Math.floor(Math.random() * maxValue)
      });
    }
    return data;
  }

  static generateSampleDataPercentage(days, minValue, maxValue) {
    const data = [];
    const today = new Date();
    for (let i = 0; i < days; i++) {
      const date = new Date(today);
      date.setDate(date.getDate() - days + i);
      data.push({
        date: format(date, 'yyyy-MM-dd'),
        value: Math.random() * (maxValue - minValue) + minValue
      });
    }
    return data;
  }

  // Convert Excel date to JS date
  static excelDateToDate(excelDate) {
    return new Date((excelDate - 25569) * 86400 * 1000);
  }

  // Parse sheet data into standardized format
  static parseSheetData(sheet, dateColumn = 0, valueColumn = 1) {
    const data = [];
    if (!sheet || sheet.length === 0) return data;

    for (let i = 1; i < sheet.length; i++) {
      const row = sheet[i];
      if (row[dateColumn] && row[valueColumn] !== undefined && row[valueColumn] !== null) {
        let date = row[dateColumn];
        if (typeof date === 'number') {
          date = this.excelDateToDate(date);
        }
        data.push({
          date: format(new Date(date), 'yyyy-MM-dd'),
          value: parseFloat(row[valueColumn]) || 0
        });
      }
    }
    return data;
  }
}

export default ExcelDataService;
