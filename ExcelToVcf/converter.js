document.getElementById('convertButton').addEventListener('click', function() 
	{
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];
    if (!file) {
        alert('请选择要转换的Excel文件哦');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(sheet);

        let vcfContent = 'BEGIN:VCARD\nVERSION:3.0\n';
        jsonData.forEach(row => {
            vcfContent += `FN:${row['FN']}\n`;
            vcfContent += `TEL;TYPE=CELL:${row['PHONE']}\n`;
            vcfContent += `ORG:${row['ORG']}\n`;
			vcfContent += `TITLE:${row['TITLE']}\n`;
            vcfContent += `END:VCARD\n`;
        });

        const blob = new Blob([vcfContent], { type: 'text/x-vcard' });
        const url = URL.createObjectURL(blob);
        const downloadLink = document.getElementById('downloadLink');
        downloadLink.href = url;
        downloadLink.download = 'contacts.vcf';
        downloadLink.style.display = 'block';
        downloadLink.textContent = '点击下载';
    };
    reader.readAsArrayBuffer(file);
});