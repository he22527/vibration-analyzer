# 背景振動量測數據分析工具

環境振動（背景振動）量測數據分析桌面應用程式與網頁版工具。

## 功能

- 讀取加速規 `.RND` 檔案（1/3 八度音帶 0.1 秒間距數據）
- 統計分析：Leq、L5、L10、L50、L90、L95、Lmax
- 單位轉換：m/s²、g、Gal、dB（加速度/速度）、μm/s、mm/s、ips
- XYZ 三軸合成
- VC 曲線比對（Ungar & Gordon 1991）：VC-E ~ 工廠 (ISO)
- 環境振動評估表
- 時間序列圖、頻譜圖、統計表

## 檔案說明

| 檔案 | 說明 |
|------|------|
| `vibration_analyzer_2.py` | 主程式（Python tkinter 桌面 App） |
| `dist/VC_Curve_Viewer.html` | 網頁版（可直接開啟或架設） |
| `make_formula_doc.py` | 產生公式說明 Excel |
| `read_sample.py` | 讀取範例數據並輸出 JSON |
| `sample_result.json` | 範例分析結果 |

## 使用方式

### 桌面版
```bash
pip install numpy pandas matplotlib
python vibration_analyzer_2.py
```

### 網頁版
直接用瀏覽器開啟 `dist/VC_Curve_Viewer.html`，或部署至任何靜態網站伺服器。

## VC 曲線標準

依據 Ungar & Gordon (1991)，評估頻段 4~80 Hz：

| 等級 | 平坦段 (μm/s) | dB (ref 2.54×10⁻⁸ m/s) |
|------|---------------|------------------------|
| 工廠 (ISO) | 803.2 | 90 |
| 辦公室 (ISO) | 402.6 | 84 |
| 住宅 (ISO) | 201.8 | 78 |
| 手術室 (ISO) | 101.1 | 72 |
| VC-A | 50.7 | 66 |
| VC-B | 25.4 | 60 |
| VC-C | 12.7 | 54 |
| VC-D | 6.38 | 48 |
| VC-E | 3.20 | 42 |

## License

MIT
