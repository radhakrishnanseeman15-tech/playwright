import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';

export default class ExcelReporter {
  constructor(options = {}) {
    this.outputFile = options.outputFile || 'playwright-test-results.xlsx';
    this.rows = [];
  }

  async onBegin(config, suite) {
    this.columns = [
      { header: 'Project', key: 'project', width: 18 },
      { header: 'File', key: 'file', width: 40 },
      { header: 'Test', key: 'test', width: 80 },
      { header: 'Outcome', key: 'outcome', width: 14 },
      { header: 'Duration ms', key: 'duration', width: 14 },
      { header: 'Errors', key: 'errors', width: 100 },
      { header: 'Start', key: 'start', width: 25 },
      { header: 'End', key: 'end', width: 25 },
    ];
  }

  async onTestEnd(test, result) {
    const errors = [];
    if (result.errors && result.errors.length) {
      for (const error of result.errors) {
        errors.push(error.message || String(error));
      }
    }

    this.rows.push({
      project: test?.project?.name || 'unknown',
      file: test?.location?.file || '',
      test: test?.title || 'unknown',
      outcome: result?.status || 'unknown',
      duration: result?.duration ?? -1,
      errors: errors.join('\n'),
      start: result?.startTime ? new Date(result.startTime).toISOString() : '',
      end: result?.endTime ? new Date(result.endTime).toISOString() : '',
    });
  }

  async onEnd(result) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Test Results');
    sheet.columns = this.columns;

    for (const row of this.rows) {
      sheet.addRow(row);
    }

    const resolvedPath = path.resolve(process.cwd(), this.outputFile);
    await workbook.xlsx.writeFile(resolvedPath);
    console.log(`Excel reporter output written to ${resolvedPath}`);
  }

  async onError(error) {
    console.error('ExcelReporter error:', error);
  }
}
