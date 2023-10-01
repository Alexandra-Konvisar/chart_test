const axios = require('axios');
const fs = require('fs');
const FormData = require('form-data');
const { Builder, Paragraph, Table, TableCell, TableRow } = require('docx');
const xl = require('excel4node');

// Создаем документ с диаграммой и таблицей
const doc = new Builder();

const data = [
  ['category', 'row 1', 'row 2', 'row 3'],
  ['category 1', '100', '100', '400'],
  ['category 2', '200', '300', '300'],
  ['category 3', '300', '400', '200'],
  ['category 4', '400', '200', '100']
];

// Создаем таблицу и заполняем данными
const table = new Table({
  rows: data.map(rowData => {
    return new TableRow({
      children: rowData.map(cellData => {
        return new TableCell({
          children: [new Paragraph(cellData)]
        });
      })
    });
  })
});

// Добавляем таблицу в документ
doc.addTable(table);

// Сохраняем документ в формате DOCX
const docxBuffer = doc.build();
fs.writeFileSync('Charty.docx', docxBuffer);

// Создаем файл XLSX с помощью библиотеки excel4node
const wb = new xl.Workbook();
const ws = wb.addWorksheet('Sheet1');

data.forEach((rowData, rowIndex) => {
  rowData.forEach((cellData, columnIndex) => {
    ws.cell(rowIndex + 1, columnIndex + 1).string(cellData);
  });
});

// Сохраняем книгу XLSX
wb.write('Charty.xlsx');


const xlsxBuilder = new Builder();
xlsxBuilder.CreateFile("xlsx");
xlsxBuilder.save(Buffer.from(fs.readFileSync('Charty.xlsx')));
xlsxBuilder.CloseFile();

const policy = {
    "expiration": "2023-12-31T12:00:00.000Z",
    "conditions": [
      {"bucket": "chartkonvisar"},
      ["starts-with", "$key", ""],
      {"acl": "public-read"},
      ["content-length-range", 0, 1048576]
    ]
  };

// Отправляем файл XLSX на сервер Amazon S3 с помощью POST-запроса
const s3BucketUrl = 'https://chartkonvisar.s3.eu-north-1.amazonaws.com/chart.js'; 
const s3AccessKey = 'AKIAX2FRS5YSFSBQISWT'; 
const s3Policy = JSON.stringify(policy);
const s3Signature = 'qavQ40UlbXs2VP7xWV9uwq61f1paR8/2mN6/vfFG'; 

const formData = new FormData();
formData.append('key', `${s3AccessKey}/Charty.xlsx`);
formData.append('AWSAccessKeyId', s3AccessKey);
formData.append('policy', s3Policy);
formData.append('signature', s3Signature);
formData.append('file', fs.createReadStream('Charty.xlsx'));

axios.post(s3BucketUrl, formData, {
  headers: {
    ...formData.getHeaders()
  }
})
  .then(response => {
    console.log('Файлы успешно отправлены на Amazon S3.');
    console.log('Ответ сервера:', response.data);
  })
  .catch(error => {
    console.error('Ошибка при отправке файлов на Amazon S3:', error);
  });
