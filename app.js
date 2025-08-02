const { createApp } = Vue;

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
            canvasHeight: 600,
            originalImageWidth: 0,
            originalImageHeight: 0,
            editScaleRatio: 1,
            isDragging: false,
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
                    const maxHeight = 600;
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
         * 添加新的文本元素
         */
        addTextElement() {
            const newElement = {
                id: this.nextElementId++,
                text: '点击编辑文本',
                x: 50,
                y: 50,
                fontSize: 24,
                color: '#000000',
                fontWeight: 'normal',
                textAlign: 'left' // 新增文本对齐属性：left, center, right
            };
            this.textElements.push(newElement);
            this.selectElement(newElement);
        },

        selectElement(element) {
            this.selectedElement = element;
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
            } else {
                // 如果没有选中的元素，则自动创建一个新的文本元素并插入变量
                const newElement = {
                    id: this.nextElementId++,
                    text: `{{${header}}}`,
                    x: 50,
                    y: 50 + (this.textElements.length * 40), // 避免重叠，每个新元素向下偏移
                    fontSize: 24,
                    color: '#000000',
                    fontWeight: 'normal',
                    textAlign: 'left' // 新增文本对齐属性
                };
                this.textElements.push(newElement);
                this.selectElement(newElement);
            }
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
                            text = text.replace(regex, rowData[header] || '');
                        });
                        
                        ctx.font = `${element.fontWeight} ${element.fontSize * finalScale}px Arial`;
                        ctx.fillStyle = element.color;
                        ctx.textBaseline = 'top';
                        
                        // 设置文本对齐方式
                        ctx.textAlign = element.textAlign || 'left';
                        
                        // 计算正确的文本位置
                        let textX = element.x * finalScale;
                        const textY = element.y * finalScale;
                        
                        // 根据对齐方式调整X坐标
                        if (element.textAlign === 'center') {
                            // 居中对齐时，X坐标保持不变，Canvas会自动居中
                        } else if (element.textAlign === 'right') {
                            // 右对齐时，X坐标保持不变，Canvas会自动右对齐
                        }
                        
                        // 处理多行文本
                        const lines = text.split('\n');
                        lines.forEach((line, lineIndex) => {
                            ctx.fillText(
                                line,
                                textX,
                                textY + lineIndex * element.fontSize * finalScale * 1.2
                            );
                        });
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
                            text = text.replace(regex, rowData[header] || '');
                        });
                        
                        ctx.font = `${element.fontWeight} ${element.fontSize * editToOriginalScale}px Arial`;
                        ctx.fillStyle = element.color;
                        ctx.textBaseline = 'top';
                        
                        // 设置文本对齐方式
                        ctx.textAlign = element.textAlign || 'left';
                        
                        // 计算正确的文本位置
                        let textX = element.x * editToOriginalScale;
                        const textY = element.y * editToOriginalScale;
                        
                        // 根据对齐方式调整X坐标
                        if (element.textAlign === 'center') {
                            // 居中对齐时，X坐标保持不变，Canvas会自动居中
                        } else if (element.textAlign === 'right') {
                            // 右对齐时，X坐标保持不变，Canvas会自动右对齐
                        }
                        
                        // 处理多行文本
                        const lines = text.split('\n');
                        lines.forEach((line, lineIndex) => {
                            ctx.fillText(
                                line,
                                textX,
                                textY + lineIndex * element.fontSize * editToOriginalScale * 1.2
                            );
                        });
                    });
                    
                    resolve(canvas);
                };
                img.src = this.backgroundImage;
            });
        }
    },

    mounted() {
        // 防止页面刷新时丢失数据，可以考虑添加localStorage支持
        console.log('批量证书生成器已加载');
    }
}).mount('#app');