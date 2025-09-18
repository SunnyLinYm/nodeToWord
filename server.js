const express = require('express');
const multer = require('multer');
const path = require('path');
const { exec } = require('child_process');
const fs = require('fs');
const sqlite3 = require('sqlite3').verbose();
const sharp = require('sharp');

const app = express();
const port = 3000;

// 設定 multer 儲存上傳的檔案
const upload = multer({ dest: 'uploads/' });

// 建立或連接 SQLite 資料庫
const db = new sqlite3.Database('data.db', (err) => {
	if (err) {
		console.error('資料庫連線錯誤:', err.message);
	} else {
		console.log('成功連線到 SQLite 資料庫。');
		// 建立表格
		db.run(`CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            text TEXT,
            image_path TEXT,
            thumbnail_path TEXT,
            remark TEXT,
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )`, (createErr) => {
			if (createErr) {
				console.error('建立表格失敗:', createErr.message);
			}
		});
	}
});

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));
app.use('/uploads/thumbnails', express.static(path.join(__dirname, 'uploads/thumbnails')));


// 處理新增資料到 Word 並存入資料庫
app.post('/add-record', upload.single('photo'), async (req, res) => {
	if (!req.file || !req.body.text) {
		return res.status(400).json({ message: '請提供文字和圖片。' });
	}

	const text = req.body.text;
	const imagePath = path.resolve(req.file.path);
	const relativeImagePath = `uploads/${path.basename(imagePath)}`;

	const thumbnailDir = path.join(__dirname, 'uploads', 'thumbnails');
	if (!fs.existsSync(thumbnailDir)) {
		fs.mkdirSync(thumbnailDir, { recursive: true });
	}
	const thumbnailPath = path.join(thumbnailDir, path.basename(imagePath));
	const relativeThumbnailPath = `uploads/thumbnails/${path.basename(imagePath)}`;

	try {
		// 使用 sharp 建立縮圖，寬度 200px
		await sharp(imagePath)
			.resize(200)
			.toFile(thumbnailPath);
	} catch (error) {
		console.error('生成縮圖失敗:', error);
		// 如果縮圖生成失敗，仍然繼續
	}
	// 呼叫 PowerShell 腳本，並傳入參數
	const psScriptPath = path.join(__dirname, 'add_to_word.ps1');
	const command = `powershell.exe -ExecutionPolicy Bypass -File "${psScriptPath}" -text "${text}" -imagePath "${imagePath}"`;

	exec(command, (error, stdout, stderr) => {
		if (error || stderr) {
			console.error(`PowerShell 執行錯誤: ${error?.message || stderr}`);
			db.run('INSERT INTO records (text, image_path, thumbnail_path, remark) VALUES (?, ?, ?, ?)',
				[text, relativeImagePath, relativeThumbnailPath, '失敗']);
			return res.status(500).json({ message: '執行 PowerShell 腳本時發生錯誤。' });
		}

		console.log(`PowerShell 執行成功: ${stdout}`);
		// 插入成功記錄到資料庫
		db.run('INSERT INTO records (text, image_path, thumbnail_path, remark) VALUES (?, ?, ?, ?)',
			[text, relativeImagePath, relativeThumbnailPath, '成功'], (dbErr) => {
				if (dbErr) {
					console.error('資料庫插入失敗:', dbErr.message);
					return res.status(500).json({ message: '資料庫插入失敗。' });
				}
				res.status(200).json({ message: '成功將資料加入到 Word 並存入資料庫！' });
			});
	});
});

// --- 新增 CRUD API 端點 ---

// 讀取所有資料 (Read)
app.get('/api/records', (req, res) => {
	db.all('SELECT * FROM records ORDER BY created_at DESC', (err, rows) => {
		if (err) {
			return res.status(500).json({ message: '讀取資料失敗。' });
		}
		res.status(200).json(rows);
	});
});

// 更新資料 (Update)
app.put('/api/records/:id', (req, res) => {
    const { id } = req.params;
    const { text, remark } = req.body;
    db.run('UPDATE records SET text = ?, remark = ? WHERE id = ?', [text, remark, id], function(err) {
        if (err) return res.status(500).json({ message: '更新資料失敗。' });
        if (this.changes === 0) return res.status(404).json({ message: '找不到該筆資料。' });
        res.status(200).json({ message: '成功更新資料。' });
    });
});

// 刪除資料 (Delete)
app.delete('/api/records/:id', (req, res) => {
	const { id } = req.params;
	db.run('DELETE FROM records WHERE id = ?', id, function (err) {
		if (err) {
			return res.status(500).json({ message: '刪除資料失敗。' });
		}
		if (this.changes === 0) {
			return res.status(404).json({ message: '找不到該筆資料。' });
		}
		res.status(200).json({ message: '成功刪除資料。' });
	});
});



app.listen(port, () => {
	console.log(`伺服器正在 http://localhost:${port} 上運行`);
	console.log('請在瀏覽器中打開此網址，或在手機上訪問此電腦的 IP 位址。');
});