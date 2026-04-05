// Danh mục khoa dựa trên mã số sinh viên
const KHOA_MAPPING = {
    "401": "Tài chính (TC)",
    "402": "Kế toán - Kiểm toán (KT-KT)",
    "403": "Quản trị kinh doanh (QTKD)",
    "404": "Công nghệ thông tin & kinh tế số (CNTT&KTS)",
    "405": "Kinh doanh quốc tế (KDQT)",
    "406": "Luật",
    "407": "Kinh tế (KT)",
    "408": "Khoa học dữ liệu (KHDL)",
    "751": "Ngôn ngữ Anh (NN)"
};

// Hàm loại bỏ dấu tiếng Việt chuẩn
function removeVietnameseTones(str) {
    if (!str) return "";
    return str
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/đ/g, 'd').replace(/Đ/g, 'D');
}

// Hàm tạo Email theo chuẩn HVNH: [tên][họ_đệm].[msv]@hvnh.edu.vn
function generateEmail(hoTen, msv) {
    if (!hoTen || !msv) return "";
    
    const cleanName = removeVietnameseTones(hoTen).toLowerCase().trim();
    const parts = cleanName.split(/\s+/);
    const msvLower = String(msv).toLowerCase().trim();

    if (parts.length === 1) {
        return `${parts[0]}.${msvLower}@hvnh.edu.vn`;
    }

    const tenChinh = parts.pop(); 
    const hoDem = parts.map(p => p.charAt(0)).join('');

    return `${tenChinh}${hoDem}.${msvLower}@hvnh.edu.vn`;
}

// Lắng nghe sự kiện tải file
document.getElementById('uploadExcel').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Lấy dữ liệu từ Sheet đầu tiên
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Chuyển sang mảng 2 chiều
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        renderToTable(rows);
    };
    reader.readAsArrayBuffer(file);
});

function renderToTable(data) {
    const tbody = document.querySelector('#studentTable tbody');
    tbody.innerHTML = ''; 

    if (data.length < 2) {
        tbody.innerHTML = '<tr><td colspan="5" class="empty-state">File không có dữ liệu phù hợp.</td></tr>';
        return;
    }

    // Tự động xác định cột dựa trên tiêu đề ở dòng 0
    let nameIdx = -1;
    let msvIdx = -1;
    const headerRow = data[0];

    headerRow.forEach((cell, idx) => {
        const val = String(cell).toLowerCase();
        if (val.includes('họ') || val.includes('tên')) nameIdx = idx;
        if (val.includes('mã') || val.includes('msv')) msvIdx = idx;
    });

    // Dự phòng nếu không tìm thấy tiêu đề: Cột 0 là MSV, Cột 1 là Tên
    if (msvIdx === -1) msvIdx = 0;
    if (nameIdx === -1) nameIdx = 1;

    // Duyệt từ dòng 1 (bỏ qua tiêu đề)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rawTen = row[nameIdx];
        const rawMsv = row[msvIdx];

        if (rawTen && rawMsv) {
            const msvStr = String(rawMsv).trim();
            const khoaHoc = msvStr.substring(0, 2);
            const maKhoa = msvStr.substring(3, 6);
            const tenKhoa = KHOA_MAPPING[maKhoa] || "Khác";
            const email = generateEmail(rawTen, msvStr);

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${rawTen}</td>
                <td><span class="msv-tag">${msvStr.toUpperCase()}</span></td>
                <td>Khóa ${khoaHoc}</td>
                <td><span class="khoa-tag">${tenKhoa}</span></td>
                <td class="email-text">${email}</td>
            `;
            tbody.appendChild(tr);
        }
    }
}