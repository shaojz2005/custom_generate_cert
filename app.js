/**
 * 批量证书生成器 - Vue.js应用
 * 
 * Copyright (c) 2024 [您的姓名或组织名称]
 * 
 * 本作品采用知识共享署名-非商业性使用 4.0 国际许可协议进行许可。
 * This work is licensed under a Creative Commons Attribution-NonCommercial 4.0 International License.
 * 
 * 您可以自由地：
 * - 共享 — 在任何媒介以任何形式复制、发行本作品
 * - 演绎 — 修改、转换或以本作品为基础进行创作
 * 
 * 惟须遵守下列条件：
 * - 署名 — 您必须给出适当的署名，提供指向本许可协议的链接，同时标明是否作了修改
 * - 非商业性使用 — 您不得将本作品用于商业目的
 * 
 * 完整许可协议：https://creativecommons.org/licenses/by-nc/4.0/legalcode.zh-Hans
 * 
 * 功能：
 * 1. Excel文件上传和解析
 * 2. 背景图片上传
 * 3. 可视化证书设计
 * 4. 批量生成和下载
 */

const { createApp } = Vue;

// 统一的字体设置常量，确保网页和Canvas使用相同字体
const FONT_FAMILY = '"Microsoft YaHei", "PingFang SC", "Hiragino Sans GB", "Helvetica Neue", Arial, sans-serif';

createApp({
    data() {
        return {
            step: 1,
            excelData: [],
            excelHeaders: [],
            backgroundImage: null,
            backgroundImageFile: null,
            textElements: [],
            selectedElement: null,
            nextElementId: 1,
            canvasWidth: 800,
            canvasHeight: 800, // 增加画布高度，提供更多设计空间
            originalImageWidth: 0,
            originalImageHeight: 0,
            editScaleRatio: 1,
            isDragging: false,
            isResizing: false,
            dragOffset: { x: 0, y: 0 },
            isGenerating: false,
            generationProgress: 0,
            currentGeneratingIndex: 0,
            showPreview: false,
            previewDataInfo: ''
        };
    },
    computed: {
        canProceedToStep2() {
            return this.excelData.length > 0 && this.backgroundImage;
        }
    },
    methods: {
        // 文件上传处理
        handleExcelUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    if (jsonData.length > 0) {
                        this.excelHeaders = jsonData[0];
                        this.excelData = jsonData.slice(1).map(row => {
                            const obj = {};
                            this.excelHeaders.forEach((header, index) => {
                                obj[header] = row[index] || '';
                            });
                            return obj;
                        });
                    }
                } catch (error) {
                    alert('Excel文件解析失败，请检查文件格式');
                }
            };
            reader.readAsArrayBuffer(file);
        },

        handleImageUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            this.backgroundImageFile = file;
            const reader = new FileReader();
            reader.onload = (e) => {
                this.backgroundImage = e.target.result;
                
                // 创建图片对象来获取原始尺寸
                const img = new Image();
                img.onload = () => {
                    // 保存原始图片尺寸
                    this.originalImageWidth = img.width;
                    this.originalImageHeight = img.height;
                    
                    // 计算合适的画布尺寸，保持宽高比
                    const maxWidth = 800;
                    const maxHeight = 800; // 增加最大高度限制
                    const ratio = Math.min(maxWidth / img.width, maxHeight / img.height);
                    
                    this.canvasWidth = Math.floor(img.width * ratio);
                    this.canvasHeight = Math.floor(img.height * ratio);
                    
                    // 保存编辑界面的缩放比例
                    this.editScaleRatio = ratio;
                };
                img.src = e.target.result;
            };
            reader.readAsDataURL(file);
        },

        // 步骤导航
        nextStep() {
            if (this.step < 3) {
                this.step++;
            }
        },

        prevStep() {
            if (this.step > 1) {
                this.step--;
            }
        },

        /**
         * 检测并格式化Excel日期数据
         * @param {any} value - 可能是日期的值
         * @returns {string} 格式化后的字符串
         */
        formatExcelValue(value) {
            // 如果值为空或undefined，返回空字符串
            if (value === null || value === undefined || value === '') {
                return '';
            }
            
            // 检测是否为Excel日期序列号（通常是5位数字，范围大约在1到50000之间）
            if (typeof value === 'number' && value > 1 && value < 100000 && Number.isInteger(value)) {
                try {
                    // Excel日期序列号转换为JavaScript日期
                    // Excel的日期起始点是1900年1月1日，但实际上Excel错误地认为1900年是闰年
                    // 所以需要减去2天来修正
                    const excelEpoch = new Date(1900, 0, 1);
                    const jsDate = new Date(excelEpoch.getTime() + (value - 2) * 24 * 60 * 60 * 1000);
                    
                    // 检查转换后的日期是否合理（1900年到2100年之间）
                    if (jsDate.getFullYear() >= 1900 && jsDate.getFullYear() <= 2100) {
                        // 格式化为 YYYY-MM-DD 格式
                        const year = jsDate.getFullYear();
                        const month = String(jsDate.getMonth() + 1).padStart(2, '0');
                        const day = String(jsDate.getDate()).padStart(2, '0');
                        return `${year}-${month}-${day}`;
                    }
                } catch (error) {
                    console.warn('日期转换失败:', value, error);
                }
            }
            
            // 如果不是日期或转换失败，返回原始值的字符串形式
            return String(value);
        },

        /**
         * 计算文本内容的自然尺寸
         * @param {string} text - 文本内容
         * @param {number} fontSize - 字体大小
         * @param {string} fontWeight - 字体粗细
         * @returns {Object} 包含width和height的对象
         */
        calculateTextSize(text, fontSize, fontWeight = 'normal', maxWidth = null) {
            // 创建临时DOM元素来测量文本尺寸，确保与HTML渲染一致
            const tempElement = document.createElement('div');
            tempElement.style.position = 'absolute';
            tempElement.style.visibility = 'hidden';
            tempElement.style.whiteSpace = 'nowrap';
            tempElement.style.fontSize = fontSize + 'px';
            tempElement.style.fontWeight = fontWeight;
            tempElement.style.fontFamily = FONT_FAMILY;
            tempElement.style.lineHeight = '1.2';
            tempElement.style.padding = '0';
            tempElement.style.margin = '0';
            tempElement.style.border = 'none';
            tempElement.textContent = text;
            
            document.body.appendChild(tempElement);
            
            // 如果没有指定最大宽度，先计算文本的自然宽度
            if (!maxWidth) {
                const naturalMaxWidth = tempElement.offsetWidth;
                
                // 检查是否是变量文本（包含双括号）
                const isVariableText = text.includes('{{') && text.includes('}}');
                
                // 设置一个合理的最大宽度限制，避免文本框过宽
                // 对于短文本和变量文本使用自然宽度，对于长文本限制最大宽度
                if (naturalMaxWidth <= 200 || isVariableText) {
                    maxWidth = naturalMaxWidth;
                } else {
                    // 长文本使用合理的最大宽度，通常是200-300px
                    maxWidth = Math.min(naturalMaxWidth, 250);
                }
            }
            
            // 设置最大宽度并允许换行来测量实际尺寸
            tempElement.style.whiteSpace = 'normal';
            tempElement.style.wordWrap = 'break-word';
            tempElement.style.wordBreak = 'break-word';
            tempElement.style.maxWidth = maxWidth + 'px';
            
            const actualWidth = tempElement.offsetWidth;
            const actualHeight = tempElement.offsetHeight;
            
            document.body.removeChild(tempElement);
            
            // 添加缓冲区确保文本有足够的显示空间
            // 浏览器的文本渲染和测量之间可能存在差异，需要更大的缓冲区
            // 对中文文本需要更大的缓冲区
            const isChinese = /[\u4e00-\u9fa5]/.test(text);
            
            // 为加粗文本增加额外的宽度冗余
            const isBold = fontWeight === 'bold' || fontWeight === '700' || fontWeight === '800' || fontWeight === '900';
            const boldExtraWidth = isBold ? Math.ceil(fontSize * 0.25) : 0; // 加粗文本额外增加25%字体大小的宽度
            
            // 为加粗文本增加更大的基础缓冲区
            const baseWidthBuffer = isBold ? 
                Math.max(16, Math.ceil(fontSize * (isChinese ? 0.5 : 0.4))) : // 加粗文本使用更大的基础缓冲区
                Math.max(12, Math.ceil(fontSize * (isChinese ? 0.4 : 0.3))); // 普通文本使用原有缓冲区
            
            const widthBuffer = baseWidthBuffer + boldExtraWidth; // 基础缓冲区 + 加粗额外宽度
            const heightBuffer = Math.max(8, Math.ceil(fontSize * (isChinese ? 0.3 : 0.2))); // 中文需要更大缓冲区
            
            return {
                width: Math.max(50, actualWidth + widthBuffer), // 增加最小宽度
                height: Math.max(fontSize * 1.5, actualHeight + heightBuffer) // 增加最小高度
            };
        },

        /**
         * 添加新的文本元素
         */
        addTextElement() {
            const text = '点击编辑文本';
            const fontSize = 16; // 增加默认字体大小
            const fontWeight = 'normal';
            
            // 对于初始文本，使用足够大的最大宽度确保不会换行
            const textSize = this.calculateTextSize(text, fontSize, fontWeight, 1000);
            
            const newElement = {
                id: this.nextElementId++,
                text: text,
                x: 50,
                y: 50,
                fontSize: fontSize,
                color: '#000000',
                fontWeight: fontWeight,
                textAlign: 'left', // 水平对齐属性：left, center, right
                verticalAlign: 'middle', // 垂直对齐属性：top, middle, bottom（默认居中）
                width: textSize.width, // 根据内容自适应宽度
                height: textSize.height // 根据内容自适应高度
            };
            this.textElements.push(newElement);
            this.selectElement(newElement);
        },

        selectElement(element) {
            console.log('selectElement called:', element);
            this.selectedElement = element;
            console.log('selectedElement set to:', this.selectedElement);
        },

        deselectElement() {
            this.selectedElement = null;
        },

        deleteSelectedElement() {
            if (this.selectedElement) {
                const index = this.textElements.indexOf(this.selectedElement);
                if (index > -1) {
                    this.textElements.splice(index, 1);
                    this.selectedElement = null;
                }
            }
        },

        /**
         * 插入Excel字段变量到文本元素
         * @param {string} header - Excel字段名
         */
        insertVariable(header) {
            // 如果有选中的元素，则在其文本后追加变量
            if (this.selectedElement) {
                this.selectedElement.text += `{{${header}}}`;
                // 重新计算尺寸
                this.updateElementSize(this.selectedElement);
            } else {
                // 如果没有选中的元素，则自动创建一个新的文本元素并插入变量
                const text = `{{${header}}}`;
                const fontSize = 16; // 增加默认字体大小
                const fontWeight = 'normal';
                
                // 对于新创建的变量文本，使用足够大的最大宽度确保不会换行
                const textSize = this.calculateTextSize(text, fontSize, fontWeight, 1000);
                
                const newElement = {
                    id: this.nextElementId++,
                    text: text,
                    x: 50,
                    y: 50 + (this.textElements.length * 40), // 避免重叠，每个新元素向下偏移
                    fontSize: fontSize,
                    color: '#000000',
                    fontWeight: fontWeight,
                    textAlign: 'left', // 水平对齐属性
                    verticalAlign: 'middle', // 垂直对齐属性（默认居中）
                    width: textSize.width, // 根据内容自适应宽度
                    height: textSize.height // 根据内容自适应高度
                };
                this.textElements.push(newElement);
                this.selectElement(newElement);
            }
        },

        /**
         * 更新文本元素的尺寸以适应内容
         * @param {Object} element - 文本元素对象
         */
        updateElementSize(element) {
            if (!element) return;
            
            // 使用当前元素的宽度作为最大宽度，如果没有设置则使用更大的默认值
            const currentWidth = element.width || 400; // 增加默认宽度
            const textSize = this.calculateTextSize(element.text, element.fontSize, element.fontWeight, currentWidth);
            
            element.width = textSize.width;
            element.height = textSize.height;
        },

        /**
         * 获取垂直对齐的CSS类
         * @param {string} verticalAlign - 垂直对齐方式
         * @returns {string} CSS flexbox align-items值
         */
        getVerticalAlignClass(verticalAlign) {
            switch (verticalAlign) {
                case 'top':
                    return 'flex-start';
                case 'middle':
                    return 'center';
                case 'bottom':
                    return 'flex-end';
                default:
                    return 'flex-start';
            }
        },

        /**
         * 获取Canvas文本基线对齐
         * @param {string} verticalAlign - 垂直对齐方式
         * @returns {string} Canvas textBaseline值
         */
        getCanvasTextBaseline(verticalAlign) {
            switch (verticalAlign) {
                case 'top':
                    return 'top';
                case 'middle':
                    return 'middle';
                case 'bottom':
                    return 'bottom';
                default:
                    return 'top';
            }
        },

        /**
         * 计算Canvas中文本的Y坐标
         * @param {Object} element - 文本元素
         * @param {number} scale - 缩放比例
         * @returns {number} 计算后的Y坐标
         */
        calculateCanvasTextY(element, scale) {
            const boxY = element.y * scale;
            const boxHeight = (element.height || 40) * scale;
            
            switch (element.verticalAlign) {
                case 'top':
                    return boxY; // 无内边距，直接从顶部开始
                case 'middle':
                    return boxY + boxHeight / 2;
                case 'bottom':
                    return boxY + boxHeight; // 无内边距，直接到底部
                default:
                    return boxY;
            }
        },

        /**
         * 获取水平对齐的CSS类
         * @param {string} textAlign - 水平对齐方式
         * @returns {string} CSS flexbox justify-content值
         */
        getHorizontalAlignClass(textAlign) {
            switch (textAlign) {
                case 'left':
                    return 'flex-start';
                case 'center':
                    return 'center';
                case 'right':
                    return 'flex-end';
                default:
                    return 'flex-start';
            }
        },

        /**
         * 文本自动换行处理
         * @param {CanvasRenderingContext2D} ctx - Canvas绘图上下文
         * @param {string} text - 要处理的文本
         * @param {number} maxWidth - 最大宽度
         * @returns {Array} 换行后的文本数组
         */
        wrapText(ctx, text, maxWidth) {
            const lines = [];
            const paragraphs = text.split('\n');
            
            paragraphs.forEach(paragraph => {
                if (paragraph.trim() === '') {
                    lines.push('');
                    return;
                }
                
                // 先按空格分词
                const words = paragraph.split(' ');
                let currentLine = '';
                
                for (let i = 0; i < words.length; i++) {
                    const word = words[i];
                    const testLine = currentLine + (currentLine ? ' ' : '') + word;
                    const metrics = ctx.measureText(testLine);
                    
                    if (metrics.width > maxWidth && currentLine !== '') {
                        lines.push(currentLine);
                        currentLine = word;
                        
                        // 检查单个单词是否仍然太长，需要强制换行
                        while (ctx.measureText(currentLine).width > maxWidth && currentLine.length > 1) {
                            let breakPoint = currentLine.length - 1;
                            
                            // 找到合适的断点
                            while (breakPoint > 0 && ctx.measureText(currentLine.substring(0, breakPoint)).width > maxWidth) {
                                breakPoint--;
                            }
                            
                            if (breakPoint === 0) breakPoint = 1; // 至少保留一个字符
                            
                            lines.push(currentLine.substring(0, breakPoint));
                            currentLine = currentLine.substring(breakPoint);
                        }
                    } else {
                        currentLine = testLine;
                    }
                }
                
                // 处理最后一行，如果仍然太长则强制换行
                while (currentLine && ctx.measureText(currentLine).width > maxWidth && currentLine.length > 1) {
                    let breakPoint = currentLine.length - 1;
                    
                    while (breakPoint > 0 && ctx.measureText(currentLine.substring(0, breakPoint)).width > maxWidth) {
                        breakPoint--;
                    }
                    
                    if (breakPoint === 0) breakPoint = 1;
                    
                    lines.push(currentLine.substring(0, breakPoint));
                    currentLine = currentLine.substring(breakPoint);
                }
                
                if (currentLine) {
                    lines.push(currentLine);
                }
            });
            
            return lines;
        },

        // 拖拽功能
        startDrag(element, event) {
            this.isDragging = true;
            this.selectElement(element);
            
            const rect = event.currentTarget.getBoundingClientRect();
            const canvasRect = event.currentTarget.parentElement.getBoundingClientRect();
            
            this.dragOffset = {
                x: event.clientX - rect.left,
                y: event.clientY - rect.top
            };

            const handleMouseMove = (e) => {
                if (this.isDragging && this.selectedElement === element) {
                    const canvasRect = document.querySelector('.canvas-container').getBoundingClientRect();
                    const newX = e.clientX - canvasRect.left - this.dragOffset.x;
                    const newY = e.clientY - canvasRect.top - this.dragOffset.y;
                    
                    element.x = Math.max(0, Math.min(newX, this.canvasWidth - 100));
                    element.y = Math.max(0, Math.min(newY, this.canvasHeight - 30));
                }
            };

            const handleMouseUp = () => {
                this.isDragging = false;
                document.removeEventListener('mousemove', handleMouseMove);
                document.removeEventListener('mouseup', handleMouseUp);
            };

            document.addEventListener('mousemove', handleMouseMove);
            document.addEventListener('mouseup', handleMouseUp);
        },

        /**
         * 开始拖拽调整尺寸
         * @param {Object} element - 文本元素对象
         * @param {string} direction - 拖拽方向：'se'(右下角), 'e'(右边), 's'(下边)
         * @param {Event} event - 鼠标事件
         */
        startResize(element, direction, event) {
            console.log('startResize called:', direction, element);
            event.preventDefault();
            event.stopPropagation();
            
            this.isResizing = true;
            this.selectElement(element);
            
            const startX = event.clientX;
            const startY = event.clientY;
            const startWidth = element.width || 200;
            const startHeight = element.height || 40;

            const handleMouseMove = (e) => {
                if (this.isResizing && this.selectedElement === element) {
                    const deltaX = e.clientX - startX;
                    const deltaY = e.clientY - startY;
                    
                    let newWidth = startWidth;
                    let newHeight = startHeight;
                    
                    // 根据拖拽方向调整尺寸
                    if (direction === 'se' || direction === 'e') {
                        newWidth = Math.max(50, startWidth + deltaX);
                    }
                    if (direction === 'se' || direction === 's') {
                        newHeight = Math.max(20, startHeight + deltaY);
                    }
                    
                    // 确保不超出画布边界
                    const maxWidth = this.canvasWidth - element.x;
                    const maxHeight = this.canvasHeight - element.y;
                    
                    element.width = Math.min(newWidth, maxWidth);
                    element.height = Math.min(newHeight, maxHeight);
                }
            };

            const handleMouseUp = () => {
                this.isResizing = false;
                document.removeEventListener('mousemove', handleMouseMove);
                document.removeEventListener('mouseup', handleMouseUp);
            };

            document.addEventListener('mousemove', handleMouseMove);
            document.addEventListener('mouseup', handleMouseUp);
        },

        // 证书生成
        async generateCertificates() {
            if (this.excelData.length === 0 || !this.backgroundImage) {
                alert('请确保已上传Excel文件和背景图片');
                return;
            }

            this.isGenerating = true;
            this.generationProgress = 0;
            this.currentGeneratingIndex = 0;

            const zip = new JSZip();
            const certificates = [];

            for (let i = 0; i < this.excelData.length; i++) {
                this.currentGeneratingIndex = i;
                this.generationProgress = Math.round((i / this.excelData.length) * 100);
                
                const rowData = this.excelData[i];
                const canvas = await this.createCertificateCanvas(rowData);
                
                // 将canvas转换为blob
                const blob = await new Promise(resolve => {
                    canvas.toBlob(resolve, 'image/png');
                });
                
                // 生成文件名，使用第一个字段的值或索引
                const fileName = rowData[this.excelHeaders[0]] || `certificate_${i + 1}`;
                zip.file(`${fileName}.png`, blob);
                
                // 添加小延迟以避免浏览器卡顿
                await new Promise(resolve => setTimeout(resolve, 10));
            }

            this.generationProgress = 100;
            
            // 生成并下载zip文件
            const zipBlob = await zip.generateAsync({ type: 'blob' });
            saveAs(zipBlob, 'certificates.zip');
            
            this.isGenerating = false;
            alert('证书生成完成！');
        },

        // 预览功能
        async previewCertificate() {
            if (this.excelData.length === 0) {
                alert('请先上传Excel文件');
                return;
            }
            
            // 使用第一条记录的数据
            const firstRowData = this.excelData[0];
            
            // 生成预览信息
            const previewInfo = this.excelHeaders.map(header => 
                `${header}: ${firstRowData[header] || ''}`
            ).join(', ');
            this.previewDataInfo = previewInfo;
            
            this.showPreview = true;
            
            // 等待DOM更新后创建预览
            await this.$nextTick();
            await this.createPreviewCanvas(firstRowData);
        },
        
        closePreview() {
            this.showPreview = false;
        },
        
        async createPreviewCanvas(rowData) {
            return new Promise((resolve) => {
                const canvas = this.$refs.previewCanvas;
                const ctx = canvas.getContext('2d');
                
                const img = new Image();
                img.onload = () => {
                    // 设置预览画布尺寸，保持合适的显示大小
                    const maxDisplayWidth = 800;
                    const maxDisplayHeight = 600;
                    const ratio = Math.min(maxDisplayWidth / img.width, maxDisplayHeight / img.height);
                    
                    canvas.width = img.width * ratio;
                    canvas.height = img.height * ratio;
                    
                    // 绘制背景图
                    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
                    
                    // 计算从编辑坐标到预览Canvas坐标的缩放比例
                    const previewScale = ratio;
                    const editToOriginalScale = 1 / this.editScaleRatio;
                    const finalScale = previewScale * editToOriginalScale;
                    
                    // 绘制文本元素
                    this.textElements.forEach(element => {
                        let text = element.text;
                        
                        // 替换变量
                        this.excelHeaders.forEach(header => {
                            const regex = new RegExp(`{{${header}}}`, 'g');
                            const formattedValue = this.formatExcelValue(rowData[header]);
                            text = text.replace(regex, formattedValue);
                        });
                        
                        ctx.font = `${element.fontWeight} ${element.fontSize * finalScale}px ${FONT_FAMILY}`;
                        ctx.fillStyle = element.color;
                        
                        // 设置文本对齐方式
                        ctx.textAlign = element.textAlign || 'left';
                        
                        // 设置垂直对齐基线
                        const verticalAlign = element.verticalAlign || 'middle';
                        ctx.textBaseline = this.getCanvasTextBaseline(verticalAlign);
                        
                        // 计算文本框的位置和尺寸
                        const boxX = element.x * finalScale;
                        const boxY = element.y * finalScale;
                        const boxWidth = (element.width || 200) * finalScale;
                        const boxHeight = (element.height || 40) * finalScale;
                        
                        // 计算文本起始位置（根据对齐方式）
                        let textX = boxX; // 左对齐时直接从左边开始
                        if (element.textAlign === 'center') {
                            textX = boxX + boxWidth / 2;
                        } else if (element.textAlign === 'right') {
                            textX = boxX + boxWidth; // 右对齐时直接到右边
                        }
                        
                        // 使用统一的Y坐标计算方法
                        let textY = this.calculateCanvasTextY(element, finalScale);
                        
                        // 保存当前绘图状态
                        ctx.save();
                        
                        // 设置裁剪区域为文本框边界
                        ctx.beginPath();
                        ctx.rect(boxX, boxY, boxWidth, boxHeight);
                        ctx.clip();
                        
                        // 处理文本自动换行
                        const lines = this.wrapText(ctx, text, boxWidth);
                        const lineHeight = element.fontSize * finalScale * 1.2;
                        
                        lines.forEach((line, lineIndex) => {
                            const currentY = textY + lineIndex * lineHeight;
                            // 只绘制在文本框高度范围内的文本行
                            if (currentY >= boxY && currentY <= boxY + boxHeight) {
                                ctx.fillText(line, textX, currentY);
                            }
                        });
                        
                        // 恢复绘图状态
                        ctx.restore();
                    });
                    
                    resolve(canvas);
                };
                img.src = this.backgroundImage;
            });
        },

        async createCertificateCanvas(rowData) {
            return new Promise((resolve) => {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                
                // 设置画布尺寸为原始图片尺寸
                const img = new Image();
                img.onload = () => {
                    canvas.width = img.width;
                    canvas.height = img.height;
                    
                    // 绘制背景图
                    ctx.drawImage(img, 0, 0);
                    
                    // 计算从编辑坐标到原始图片坐标的缩放比例
                    const editToOriginalScale = 1 / this.editScaleRatio;
                    
                    // 绘制文本元素
                    this.textElements.forEach(element => {
                        let text = element.text;
                        
                        // 替换变量
                        this.excelHeaders.forEach(header => {
                            const regex = new RegExp(`{{${header}}}`, 'g');
                            const formattedValue = this.formatExcelValue(rowData[header]);
                            text = text.replace(regex, formattedValue);
                        });
                        
                        ctx.font = `${element.fontWeight} ${element.fontSize * editToOriginalScale}px ${FONT_FAMILY}`;
                        ctx.fillStyle = element.color;
                        
                        // 设置文本对齐方式
                        ctx.textAlign = element.textAlign || 'left';
                        
                        // 设置垂直对齐基线
                        const verticalAlign = element.verticalAlign || 'middle';
                        ctx.textBaseline = this.getCanvasTextBaseline(verticalAlign);
                        
                        // 计算文本框的位置和尺寸
                        const boxX = element.x * editToOriginalScale;
                        const boxY = element.y * editToOriginalScale;
                        const boxWidth = (element.width || 200) * editToOriginalScale;
                        const boxHeight = (element.height || 40) * editToOriginalScale;
                        
                        // 计算文本起始位置（根据对齐方式）
                        let textX = boxX; // 左对齐时直接从左边开始
                        if (element.textAlign === 'center') {
                            textX = boxX + boxWidth / 2;
                        } else if (element.textAlign === 'right') {
                            textX = boxX + boxWidth; // 右对齐时直接到右边
                        }
                        
                        // 使用统一的Y坐标计算方法
                        let textY = this.calculateCanvasTextY(element, editToOriginalScale);
                        
                        // 保存当前绘图状态
                        ctx.save();
                        
                        // 设置裁剪区域为文本框边界
                        ctx.beginPath();
                        ctx.rect(boxX, boxY, boxWidth, boxHeight);
                        ctx.clip();
                        
                        // 处理文本自动换行
                        const lines = this.wrapText(ctx, text, boxWidth);
                        const lineHeight = element.fontSize * editToOriginalScale * 1.2;
                        
                        lines.forEach((line, lineIndex) => {
                            const currentY = textY + lineIndex * lineHeight;
                            // 只绘制在文本框高度范围内的文本行
                            if (currentY >= boxY && currentY <= boxY + boxHeight) {
                                ctx.fillText(line, textX, currentY);
                            }
                        });
                        
                        // 恢复绘图状态
                        ctx.restore();
                    });
                    
                    resolve(canvas);
                };
                img.src = this.backgroundImage;
            });
        },

        /**
         * 格式化文本以在HTML中显示，处理换行符
         * @param {string} text - 要格式化的文本
         * @returns {string} 格式化后的HTML字符串
         */
        formatTextForDisplay(text) {
            if (!text) return '';
            
            // 转义HTML特殊字符
            const escapeHtml = (str) => {
                const div = document.createElement('div');
                div.textContent = str;
                return div.innerHTML;
            };
            
            // 转义HTML并将换行符转换为<br>标签
            return escapeHtml(text).replace(/\n/g, '<br>');
        }
    },

    mounted() {
        // 防止页面刷新时丢失数据，可以考虑添加localStorage支持
        console.log('批量证书生成器已加载');
    }
}).mount('#app');