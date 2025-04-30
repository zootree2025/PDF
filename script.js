document.addEventListener('DOMContentLoaded', function() {
    const fontColor = document.getElementById('fontColor');
    const bgColor = document.getElementById('bgColor');
    const fontColorPreview = document.getElementById('fontColorPreview');
    const bgColorPreview = document.getElementById('bgColorPreview');

    // 即時預覽顏色選擇
    fontColor.addEventListener('input', function () {
        fontColorPreview.style.backgroundColor = this.value;
    });

    bgColor.addEventListener('input', function () {
        bgColorPreview.style.backgroundColor = this.value;
    });

    // 初始化顏色預覽
    fontColorPreview.style.backgroundColor = fontColor.value;
    bgColorPreview.style.backgroundColor = bgColor.value;

    // 轉換按鈕點擊事件
    document.getElementById('convertBtn').addEventListener('click', function () {
        const fileInput = document.getElementById('txtFile');
        const file = fileInput.files[0];

        if (!file) {
            alert('請選擇一個 TXT 文件');
            return;
        }

        const reader = new FileReader();
        reader.onload = function (e) {
            let content = e.target.result;
            content = content.replace(/\r\n/g, '\n'); // 統一換行符為 Unix 格式
            createPPT(content, fontColor.value, bgColor.value);
        };
        reader.readAsText(file, 'UTF-8');
    });

    // 核心 PPT 生成邏輯
    function createPPT(content, fontColor, bgColor) {
        try {
            if (typeof PptxGenJS === 'undefined') {
                throw new Error('PptxGenJS 库未正確加載，請刷新頁面重試');
            }

            const pptx = new PptxGenJS();
            pptx.defineLayout({ name: 'LAYOUT_16x9', width: 10, height: 5.625 });
            pptx.layout = 'LAYOUT_16x9';

            // 使用修正後的正則表達式分割投影片
            const slides = content.split(/\n\s*\n/); // 關鍵修正點

            slides.forEach(slideContent => {
                if (slideContent.trim() === '') return; // 跳過空白內容

                const slide = pptx.addSlide();
                slide.background = { color: bgColor };

                // 處理 /N 換行符號
                const processedContent = slideContent.replace(/\/N/g, '\n');

                slide.addText(processedContent, {
                    x: 0.5,
                    y: 0.5,
                    w: '90%',
                    h: '80%',
                    color: fontColor,
                    fontSize: 24,
                    align: 'left',
                    valign: 'top',
                    bold: false // 可選：是否加粗文字
                });
            });

            // 生成並提供下載
            pptx.writeFile({ fileName: 'presentation.pptx' })
                .then(() => {
                    document.getElementById('result').classList.remove('hidden');
                })
                .catch(err => {
                    console.error('PPT 生成失敗:', err);
                    alert('PPT 生成失敗: ' + err.message);
                });

        } catch (error) {
            console.error('生成 PPT 時發生錯誤:', error);
            alert('生成 PPT 時發生錯誤: ' + error.message);
        }
    }
});