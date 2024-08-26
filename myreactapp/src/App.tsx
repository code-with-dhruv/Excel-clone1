import React, { useState, useEffect, useCallback, useRef } from 'react';
import { SpreadsheetComponent, SheetModel } from '@syncfusion/ej2-react-spreadsheet';
import { UploaderComponent } from '@syncfusion/ej2-react-inputs';
import './App.css';
import * as XLSX from 'xlsx';
import { registerLicense } from '@syncfusion/ej2-base';
import ReactDOM from 'react-dom';

// Register the Syncfusion license key
registerLicense('Ngo9BigBOggjHTQxAR8/V1NCaF1cWWhAYVF1WmFZfVpgd19FaVZVRmYuP1ZhSXxXdk1hXX9dc31RRWVYU0I=');

interface SheetData {
  [key: string]: string | number;
}

function App() {
  const spreadsheetRef = useRef<SpreadsheetComponent>(null);

  // Function to handle file upload and load it into the Spreadsheet
  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = e.target?.result;
        if (data) {
          const workbook = XLSX.read(data, { type: 'binary' });

          const sheetsData: SheetData[][] = workbook.SheetNames.map(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            return XLSX.utils.sheet_to_json(sheet, { header: 1 });
          });

          if (spreadsheetRef.current) {
            const sheets: SheetModel[] = workbook.SheetNames.map((sheetName, index) => {
              const sheet = sheetsData[index];
              const colCount = sheet[0] ? Object.keys(sheet[0]).length : 0;

              return {
                name: sheetName,
                ranges: [{ dataSource: sheet, startCell: 'A1' }],
                rowCount: sheet.length,
                colCount: colCount,
              };
            });

            spreadsheetRef.current.sheets = sheets;
            spreadsheetRef.current.refresh();
          }
        }
      };
      reader.readAsBinaryString(file);
    }
  };

  // Fetch and load a remote Excel file on component mount
  useEffect(() => {
    const fetchData = async (): Promise<void> => {
      const response: Response = await fetch(
        'https://cdn.syncfusion.com/scripts/spreadsheet/Sample.xlsx'
      );
      const fileBlob: Blob = await response.blob();
      const file: File = new File([fileBlob], 'Sample.xlsx');
      const spreadsheet = spreadsheetRef.current;
      if (spreadsheet) {
        spreadsheet.open({ file });
      }
    };
    fetchData();
  }, []);

  return (
    <div className="App" style={{ height: '100vh', overflowY: 'auto' }}>
      <input type="file" accept=".xlsm" onChange={handleFileUpload} />
      <SpreadsheetComponent
        ref={spreadsheetRef}
        allowScrolling={true}
        scrollSettings={{
          enableVirtualization: true,
        }}
        height={'100rem'}
        width={'100vw'}
        enablePersistence={false}
        allowDataValidation={false}
        openUrl="https://services.syncfusion.com/react/production/api/spreadsheet/open"
      />
    </div>
  );
}

export default React.memo(App);

ReactDOM.render(<App />, document.getElementById('root'));
