import axios from 'axios';
import ExcelJS from 'exceljs';

async function verifyApi() {
  const apiUrl = 'http://localhost:4000/api/generate/excel';
  console.log(`Checking API at ${apiUrl}...`);

  // 1. Create a dummy template
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Test Sheet');
  worksheet.getCell('A1').value = 'Hello from API';
  worksheet.getCell('B1').value = '{{ placeholder }}';

  const buffer = await workbook.xlsx.writeBuffer();
  const templateBase64 = Buffer.from(buffer).toString('base64');

  // 2. Send Request
  try {
    const response = await axios.post(apiUrl, {
      templateBase64,
      data: { placeholder: 'Substituted Value' }
    });

    if (response.status === 200 && response.data.success && response.data.data) {
      console.log('✅ API Request Successful');
      console.log('MimeType:', response.data.mimeType);

      // Optional: decode data and check content
      const resultBuffer = Buffer.from(response.data.data, 'base64');
      const resultWorkbook = new ExcelJS.Workbook();
      await resultWorkbook.xlsx.load(resultBuffer as any);
      const resultSheet = resultWorkbook.getWorksheet('Test Sheet');
      const cellValue = resultSheet?.getCell('B1').value;

      if (cellValue === 'Substituted Value') {
        console.log('✅ Placeholder substitution verified');
      } else {
        console.error(`❌ Substitution failed, got: ${cellValue}`);
      }

    } else {
      console.error('❌ API Request Failed:', response.data);
    }

  } catch (error: any) {
    if (axios.isAxiosError(error)) {
      console.error('❌ API Error:', error.message);
      if (error.response) {
        console.error('Data:', error.response.data);
      }
    } else {
      console.error('❌ Unexpected Error:', error);
    }
    process.exit(1);
  }
}

verifyApi();
