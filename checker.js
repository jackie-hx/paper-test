const FormatChecker = {
    rules: {
        title: {
            name: '论文标题',
            fontSize: 22,
            fontWeight: 'bold',
            alignment: 'center'
        },
        subtitle: {
            name: '副标题',
            fontSize: 18,
            fontWeight: 'bold'
        },
        chapterTitle: {
            name: '一级标题（章）',
            fontSize: 16,
            fontWeight: 'bold'
        },
        sectionTitle: {
            name: '二级标题（节）',
            fontSize: 14,
            fontWeight: 'bold'
        },
        subsectionTitle: {
            name: '三级标题',
            fontSize: 12,
            fontWeight: 'bold'
        },
        body: {
            name: '正文',
            fontSize: 12,
            lineHeight: 1.5,
            firstLineIndent: 2,
            fontFamily: '宋体'
        },
        abstract: {
            name: '摘要',
            fontSize: 12,
            minWords: 200,
            maxWords: 500
        },
        keywords: {
            name: '关键词',
            minCount: 3,
            maxCount: 5
        },
        reference: {
            name: '参考文献',
            minCount: 10
        }
    },

    chapterPatterns: [
        /^第[一二三四五六七八九十百]+章/,
        /^第[一二三四五六七八九十百]+部分/,
        /^Chapter\s*\d+/i,
        /^\d+[\s\.、]+.{2,20}$/,
        /^[一二三四五六七八九十百]+[\s\.、]+.{2,20}$/
    ],

    specialChapterNames: [
        '绪论', '引言', '研究背景', '文献综述', '理论基础', 
        '研究方法', '研究设计', '实证分析', '案例分析',
        '研究结果', '结果分析', '讨论', '结论', '总结',
        '参考文献', '致谢', '附录', '摘要', 'Abstract',
        '目录', '引言', '正文', '结语'
    ],

    checkDocument(parsedDoc) {
        const results = {
            formatIssues: [],
            contentIssues: [],
            structureIssues: [],
            structureInfo: {},
            summary: {
                totalIssues: 0,
                errors: 0,
                warnings: 0,
                infos: 0
            }
        };

        this.detectDocumentStructure(parsedDoc, results);
        this.checkStructure(parsedDoc, results);
        this.checkFormat(parsedDoc, results);
        this.checkContent(parsedDoc, results);
        this.calculateSummary(results);

        return results;
    },

    detectDocumentStructure(doc, results) {
        const structure = {
            hasTitle: false,
            hasChineseAbstract: false,
            hasEnglishAbstract: false,
            hasKeywords: false,
            hasEnglishKeywords: false,
            hasTOC: false,
            hasIntroduction: false,
            hasConclusion: false,
            hasReferences: false,
            hasAcknowledgment: false,
            hasAppendix: false,
            chapterCount: 0,
            chapters: [],
            wordCount: 0,
            paragraphCount: doc.paragraphs ? doc.paragraphs.length : 0
        };

        const text = doc.text || '';
        structure.wordCount = text.replace(/\s/g, '').length;

        if (doc.paragraphs) {
            let foundChapters = [];
            
            doc.paragraphs.forEach((para, index) => {
                const paraText = (para.text || para).trim();
                const cleanText = paraText.replace(/\s+/g, '');
                
                if (this.isPaperTitle(paraText, index, doc.paragraphs)) {
                    structure.hasTitle = true;
                }
                
                if (this.isChineseAbstractTitle(paraText)) {
                    structure.hasChineseAbstract = true;
                }
                
                if (this.isEnglishAbstractTitle(paraText)) {
                    structure.hasEnglishAbstract = true;
                }
                
                if (this.isChineseKeywords(paraText)) {
                    structure.hasKeywords = true;
                }
                
                if (this.isEnglishKeywords(paraText)) {
                    structure.hasEnglishKeywords = true;
                }
                
                if (this.isTOC(paraText)) {
                    structure.hasTOC = true;
                }
                
                if (this.isReferences(paraText)) {
                    structure.hasReferences = true;
                }
                
                if (this.isAcknowledgment(paraText)) {
                    structure.hasAcknowledgment = true;
                }
                
                if (this.isAppendix(paraText)) {
                    structure.hasAppendix = true;
                }
                
                const chapterInfo = this.detectChapter(paraText, index);
                if (chapterInfo) {
                    foundChapters.push({
                        name: chapterInfo.name,
                        index: index,
                        type: chapterInfo.type
                    });
                    
                    if (this.isIntroductionChapter(paraText)) {
                        structure.hasIntroduction = true;
                    }
                    if (this.isConclusionChapter(paraText)) {
                        structure.hasConclusion = true;
                    }
                }
            });
            
            structure.chapters = foundChapters;
            structure.chapterCount = foundChapters.filter(c => 
                c.type === 'numbered' || c.type === 'special'
            ).length;
        }

        results.structureInfo = structure;
    },

    detectChapter(text, index) {
        const trimmed = text.trim();
        
        for (let pattern of this.chapterPatterns) {
            if (pattern.test(trimmed)) {
                return {
                    name: trimmed,
                    type: 'numbered'
                };
            }
        }
        
        for (let name of this.specialChapterNames) {
            if (trimmed === name || trimmed === name + '：' || trimmed === name + ':') {
                return {
                    name: trimmed,
                    type: 'special'
                };
            }
        }
        
        if (trimmed.length >= 2 && trimmed.length <= 15) {
            const hasChineseOnly = /^[\u4e00-\u9fa5]+$/.test(trimmed);
            const isLikelyTitle = hasChineseOnly && !this.isBodyContent(trimmed);
            if (isLikelyTitle) {
                return {
                    name: trimmed,
                    type: 'potential'
                };
            }
        }
        
        return null;
    },

    isBodyContent(text) {
        const bodyIndicators = ['的', '了', '是', '在', '有', '和', '与', '或', '等', '及'];
        return bodyIndicators.some(indicator => text.includes(indicator));
    },

    isPaperTitle(text, index, paragraphs) {
        if (index > 5) return false;
        const trimmed = text.trim();
        if (trimmed.length < 5 || trimmed.length > 60) return false;
        
        const excludePatterns = [
            /^摘要/, /^Abstract/i, /^关键词/, /^Keywords/i,
            /^目录/, /^参考文献/, /^致谢/, /^附录/,
            /^第[一二三四五六七八九十]+章/, /^\d+[\.\s]/
        ];
        
        for (let pattern of excludePatterns) {
            if (pattern.test(trimmed)) return false;
        }
        
        return true;
    },

    isChineseAbstractTitle(text) {
        const trimmed = text.trim();
        return /^摘\s*要$/.test(trimmed) || 
               trimmed === '摘要' || 
               trimmed === '中文摘要' ||
               /^【摘要】$/.test(trimmed);
    },

    isEnglishAbstractTitle(text) {
        const trimmed = text.trim();
        return /^Abstract$/i.test(trimmed) || 
               trimmed.toLowerCase() === 'abstract';
    },

    isChineseKeywords(text) {
        const trimmed = text.trim();
        return /^关键词\s*[:：]/.test(trimmed) ||
               /^【关键词】/.test(trimmed) ||
               /^关键词$/.test(trimmed);
    },

    isEnglishKeywords(text) {
        const trimmed = text.trim();
        return /^Keywords\s*[:：]/i.test(trimmed) ||
               /^Keywords$/i.test(trimmed);
    },

    isIntroductionChapter(text) {
        const trimmed = text.trim();
        return /绪论|引言|研究背景/.test(trimmed) ||
               /^第[一二三四五六七八九十]+章\s*.*绪论/.test(trimmed) ||
               /^1[\s\.、]+.*绪论/.test(trimmed);
    },

    isConclusionChapter(text) {
        const trimmed = text.trim();
        return /结论|总结|结语/.test(trimmed) ||
               /^第[一二三四五六七八九十]+章\s*.*结论/.test(trimmed);
    },

    checkStructure(doc, results) {
        const structure = results.structureInfo;

        if (!structure.hasTitle) {
            this.addIssue(results.structureIssues, 'error', '缺少论文标题', '论文应包含标题，标题应居中显示，字号一般为小二号或三号', '文档开头');
        }

        if (!structure.hasChineseAbstract) {
            this.addIssue(results.structureIssues, 'error', '缺少中文摘要', '论文应包含中文摘要，摘要标题应为"摘要"，内容应在200-500字之间', '正文前');
        }

        if (!structure.hasEnglishAbstract) {
            this.addIssue(results.structureIssues, 'warning', '缺少英文摘要', '建议添加英文摘要（Abstract），便于国际交流', '中文摘要后');
        }

        if (!structure.hasKeywords) {
            this.addIssue(results.structureIssues, 'warning', '缺少中文关键词', '论文应包含3-5个中文关键词，格式为"关键词：关键词1；关键词2；..."', '摘要之后');
        }

        if (!structure.hasEnglishKeywords && structure.hasEnglishAbstract) {
            this.addIssue(results.structureIssues, 'info', '缺少英文关键词', '建议在英文摘要后添加英文关键词（Keywords）', '英文摘要后');
        }

        if (!structure.hasTOC) {
            this.addIssue(results.structureIssues, 'info', '未检测到目录', '建议在摘要后添加目录，便于读者快速定位内容', '摘要后');
        }

        if (!structure.hasIntroduction) {
            this.addIssue(results.structureIssues, 'warning', '缺少绪论/引言章节', '论文应包含绪论或引言章节，介绍研究背景和目的', '正文开头');
        }

        if (!structure.hasConclusion) {
            this.addIssue(results.structureIssues, 'warning', '缺少结论章节', '论文应包含结论章节，总结研究成果和贡献', '正文末尾');
        }

        if (!structure.hasReferences) {
            this.addIssue(results.structureIssues, 'error', '缺少参考文献', '论文应包含参考文献列表，建议不少于10篇，格式需符合规范', '论文末尾');
        }

        if (!structure.hasAcknowledgment) {
            this.addIssue(results.structureIssues, 'info', '缺少致谢', '建议在参考文献后添加致谢，感谢指导老师和帮助过的人', '参考文献后');
        }

        if (structure.chapterCount < 3) {
            this.addIssue(results.structureIssues, 'warning', '章节结构不完整', `当前检测到${structure.chapterCount}个主要章节，建议论文包含：绪论、正文（多个章节）、结论等完整结构`, '全文');
        }

        if (structure.chapters.length > 0) {
            this.checkChapterOrder(structure.chapters, results);
        }

        if (structure.wordCount < 5000) {
            this.addIssue(results.structureIssues, 'warning', '论文字数偏少', `当前字数约${structure.wordCount}字，本科毕业论文一般要求8000字以上`, '全文');
        } else if (structure.wordCount < 8000) {
            this.addIssue(results.structureIssues, 'info', '论文字数提醒', `当前字数约${structure.wordCount}字，建议达到8000字以上`, '全文');
        }
    },

    checkChapterOrder(chapters, results) {
        const mainChapters = chapters.filter(c => c.type === 'numbered' || c.type === 'special');
        
        if (mainChapters.length > 0) {
            let lastNumberedIndex = -1;
            for (let chapter of mainChapters) {
                if (chapter.type === 'numbered') {
                    const match = chapter.name.match(/(\d+)/);
                    if (match) {
                        const num = parseInt(match[1]);
                        if (num < lastNumberedIndex) {
                            this.addIssue(results.structureIssues, 'warning', '章节编号顺序异常', `检测到章节"${chapter.name}"编号可能存在顺序问题`, `第${chapter.index + 1}段`);
                        }
                        lastNumberedIndex = num;
                    }
                }
            }
        }
    },

    checkFormat(doc, results) {
        if (!doc.paragraphs) return;

        let inAbstractSection = false;
        let abstractContent = '';
        let inReferenceSection = false;

        doc.paragraphs.forEach((para, index) => {
            const text = (para.text || para).trim();
            
            if (this.isChineseAbstractTitle(text) || this.isEnglishAbstractTitle(text)) {
                inAbstractSection = true;
                abstractContent = '';
                return;
            }
            
            if (inAbstractSection && text.length > 0) {
                if (this.isKeywords(text) || this.isEnglishKeywords(text) || 
                    this.isTOC(text) || this.detectChapter(text, index)) {
                    if (abstractContent.length > 0) {
                        this.checkAbstractContent(abstractContent, results);
                    }
                    inAbstractSection = false;
                    abstractContent = '';
                } else {
                    abstractContent += text;
                }
            }
            
            if (this.isReferences(text)) {
                inReferenceSection = true;
            }
            
            if (this.detectChapter(text, index)) {
                this.checkChapterTitleFormat(text, index, results);
            }
            
            if (this.isSectionTitle(text)) {
                this.checkSectionTitleFormat(text, index, results);
            }
            
            if (this.isBodyParagraph(text) && !inReferenceSection) {
                this.checkBodyFormat(text, index, results);
            }
            
            if (inReferenceSection && this.isReferenceItem(text)) {
                this.checkReferenceFormat(text, index, results);
            }
        });

        this.checkImageTableFormat(doc, results);
    },

    checkAbstractContent(content, results) {
        const wordCount = content.replace(/\s/g, '').length;
        if (wordCount < 200) {
            this.addIssue(results.formatIssues, 'warning', '摘要字数不足', `摘要内容约${wordCount}字，建议在200-500字之间`, '摘要部分');
        } else if (wordCount > 500) {
            this.addIssue(results.formatIssues, 'warning', '摘要字数过多', `摘要内容约${wordCount}字，建议控制在500字以内`, '摘要部分');
        }
    },

    checkContent(doc, results) {
        const text = doc.text || '';
        
        const emptyParagraphs = this.countEmptyParagraphs(doc);
        if (emptyParagraphs > 5) {
            this.addIssue(results.formatIssues, 'info', '存在多个空段落', `检测到${emptyParagraphs}个空段落，建议删除多余空行`, '全文');
        }

        const longParagraphs = this.findLongParagraphs(doc);
        longParagraphs.forEach(p => {
            this.addIssue(results.formatIssues, 'warning', '段落过长', `第${p.index + 1}段文字较长，建议适当分段以提高可读性`, `第${p.index + 1}段`);
        });

        const repeatedWords = this.checkRepeatedWords(text);
        if (repeatedWords.length > 0) {
            this.addIssue(results.contentIssues, 'info', '存在重复词语', `检测到可能的重复词语：${repeatedWords.slice(0, 3).join('、')}`, '全文');
        }
    },

    isTitle(text, index) {
        return index < 5 && text.length > 5 && text.length < 50 && !text.includes('：') && !text.includes(':');
    },

    isAbstract(text) {
        return /^摘要|^【摘要】|^\s*摘要/.test(text);
    },

    isKeywords(text) {
        return this.isChineseKeywords(text) || this.isEnglishKeywords(text);
    },

    isTOC(text) {
        const trimmed = text.trim();
        return /^目录$/.test(trimmed) || /^【目录】$/.test(trimmed);
    },

    isAppendix(text) {
        const trimmed = text.trim();
        return /^附录$/.test(trimmed) || /^【附录】$/.test(trimmed) || /^附录[一二三四五六七八九十]/.test(trimmed);
    },

    isReferences(text) {
        const trimmed = text.trim();
        return /^参考文献$/.test(trimmed) || /^【参考文献】$/.test(trimmed) || /^References$/i.test(trimmed);
    },

    isAcknowledgment(text) {
        const trimmed = text.trim();
        return /^致谢$/.test(trimmed) || /^【致谢】$/.test(trimmed);
    },

    isChapterTitle(text) {
        return this.detectChapter(text, 0) !== null;
    },

    isSectionTitle(text) {
        const trimmed = text.trim();
        return /^\d+\.\d+/.test(trimmed) && trimmed.length < 50;
    },

    isBodyParagraph(text) {
        const trimmed = text.trim();
        return trimmed.length > 50 && 
               !this.isChapterTitle(trimmed) && 
               !this.isSectionTitle(trimmed) && 
               !this.isReferences(trimmed) &&
               !this.isChineseAbstractTitle(trimmed) &&
               !this.isEnglishAbstractTitle(trimmed);
    },

    isReferenceItem(text) {
        const trimmed = text.trim();
        return /^\[\d+\]|^\d+[\.\、]/.test(trimmed) && trimmed.length > 20;
    },

    checkChapterTitleFormat(text, index, results) {
        const trimmed = text.trim();
        
        if (/^第[一二三四五六七八九十]+章/.test(trimmed)) {
            if (!/^第[一二三四五六七八九十]+章\s+.{2,}/.test(trimmed)) {
                this.addIssue(results.formatIssues, 'warning', '章节标题格式建议', '章节标题建议使用"第X章 标题"格式，章节号与标题之间应有空格', `第${index + 1}段`);
            }
        }
    },

    checkSectionTitleFormat(text, index, results) {
        const trimmed = text.trim();
        const match = trimmed.match(/^(\d+)\.(\d+)/);
        if (match) {
            const sectionNum = parseInt(match[2]);
            if (sectionNum > 20) {
                this.addIssue(results.formatIssues, 'warning', '节编号过大', `检测到节编号"${match[0]}"，请确认编号是否正确`, `第${index + 1}段`);
            }
        }
    },

    checkBodyFormat(text, index, results) {
        const trimmed = text.trim();
        
        if (trimmed.length > 0 && !/^\s/.test(text)) {
            this.addIssue(results.formatIssues, 'info', '段落首行未缩进', '正文段落首行应缩进两个字符', `第${index + 1}段`);
        }

        if (trimmed.length > 500) {
            this.addIssue(results.formatIssues, 'warning', '段落过长', '单个段落超过500字，建议适当分段', `第${index + 1}段`);
        }
    },

    checkReferenceFormat(text, index, results) {
        const trimmed = text.trim();
        
        if (!/\[\d+\]/.test(trimmed) && !/^\d+[\.\、]/.test(trimmed)) {
            this.addIssue(results.formatIssues, 'info', '参考文献格式不规范', '参考文献建议使用"［1］"或"[1]"格式编号', `第${index + 1}段`);
        }

        if (!/[期刊|J]|书籍|M|学位论文|D|会议|C|报纸|N|报告|R|标准|S|专利|P/.test(trimmed) && !trimmed.includes('http')) {
            this.addIssue(results.formatIssues, 'info', '参考文献缺少文献类型标识', '参考文献应标注文献类型，如[J]、[M]、[D]等', `第${index + 1}段`);
        }
    },

    checkImageTableFormat(doc, results) {
        const text = doc.text || '';
        
        const figurePattern = /图\s*\d+/g;
        const tablePattern = /表\s*\d+/g;
        const figures = text.match(figurePattern) || [];
        const tables = text.match(tablePattern) || [];

        if (figures.length > 0) {
            const figureCaptionPattern = /图\s*\d+[\s\-:：]/;
            figures.forEach((f, i) => {
                if (!figureCaptionPattern.test(text)) {
                    this.addIssue(results.formatIssues, 'info', '图片标题格式不规范', '图片标题应使用"图X-X 标题"格式', '全文');
                }
            });
        }

        if (tables.length > 0) {
            const tableCaptionPattern = /表\s*\d+[\s\-:：]/;
            tables.forEach((t, i) => {
                if (!tableCaptionPattern.test(text)) {
                    this.addIssue(results.formatIssues, 'info', '表格标题格式不规范', '表格标题应使用"表X-X 标题"格式', '全文');
                }
            });
        }
    },

    countEmptyParagraphs(doc) {
        if (!doc.paragraphs) return 0;
        return doc.paragraphs.filter(p => {
            const text = p.text || p;
            return text.trim().length === 0;
        }).length;
    },

    findLongParagraphs(doc) {
        if (!doc.paragraphs) return [];
        return doc.paragraphs.map((p, i) => {
            const text = p.text || p;
            return { index: i, length: text.length };
        }).filter(p => p.length > 600);
    },

    checkRepeatedWords(text) {
        const words = text.match(/[\u4e00-\u9fa5]{2,4}/g) || [];
        const wordCount = {};
        words.forEach(w => {
            wordCount[w] = (wordCount[w] || 0) + 1;
        });
        return Object.entries(wordCount)
            .filter(([w, c]) => c > 10 && !['可以', '进行', '通过', '根据', '以及', '但是', '因此', '所以', '因为', '如果'].includes(w))
            .map(([w, c]) => w);
    },

    addIssue(list, type, title, description, location) {
        list.push({
            type,
            title,
            description,
            location
        });
    },

    calculateSummary(results) {
        const countByType = (list, type) => list.filter(i => i.type === type).length;
        
        results.summary.errors = countByType(results.formatIssues, 'error') + 
                                  countByType(results.contentIssues, 'error') + 
                                  countByType(results.structureIssues, 'error');
        results.summary.warnings = countByType(results.formatIssues, 'warning') + 
                                    countByType(results.contentIssues, 'warning') + 
                                    countByType(results.structureIssues, 'warning');
        results.summary.infos = countByType(results.formatIssues, 'info') + 
                                 countByType(results.contentIssues, 'info') + 
                                 countByType(results.structureIssues, 'info');
        results.summary.totalIssues = results.summary.errors + results.summary.warnings + results.summary.infos;
    }
};

const PaperChecker = {
    currentFile: null,
    parsedDocument: null,
    checkResults: null,

    init() {
        this.bindEvents();
    },

    bindEvents() {
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const checkBtn = document.getElementById('checkBtn');
        const clearBtn = document.getElementById('clearBtn');
        const exportBtn = document.getElementById('exportBtn');

        uploadArea.addEventListener('click', () => fileInput.click());
        
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.handleFile(files[0]);
            }
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.handleFile(e.target.files[0]);
            }
        });

        checkBtn.addEventListener('click', () => this.runCheck());
        clearBtn.addEventListener('click', () => this.clearFile());
        exportBtn.addEventListener('click', () => this.exportToExcel());

        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => this.switchTab(e.target.dataset.tab));
        });
    },

    handleFile(file) {
        if (!file.name.endsWith('.docx')) {
            alert('请上传 .docx 格式的文件！');
            return;
        }

        this.currentFile = file;
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('fileInfo').style.display = 'flex';
        document.getElementById('uploadArea').style.display = 'none';
        document.getElementById('checkBtn').disabled = false;
    },

    clearFile() {
        this.currentFile = null;
        this.parsedDocument = null;
        this.checkResults = null;
        document.getElementById('fileInput').value = '';
        document.getElementById('fileInfo').style.display = 'none';
        document.getElementById('uploadArea').style.display = 'block';
        document.getElementById('checkBtn').disabled = true;
        document.getElementById('resultSection').style.display = 'none';
    },

    async runCheck() {
        if (!this.currentFile) return;

        const checkBtn = document.getElementById('checkBtn');
        const btnText = checkBtn.querySelector('.btn-text');
        const btnLoading = checkBtn.querySelector('.btn-loading');

        checkBtn.disabled = true;
        btnText.style.display = 'none';
        btnLoading.style.display = 'flex';

        try {
            const arrayBuffer = await this.readFileAsArrayBuffer(this.currentFile);
            
            if (typeof mammoth === 'undefined') {
                throw new Error('mammoth.js 库未正确加载');
            }
            
            const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
            
            if (!result || !result.value) {
                throw new Error('文件内容解析失败');
            }

            const paragraphs = result.value.split(/\r?\n/).filter(p => p.trim().length > 0);
            
            this.parsedDocument = {
                text: result.value,
                paragraphs: paragraphs.map(p => ({ text: p }))
            };

            const checkFormat = document.getElementById('checkFormat').checked;
            const checkStructure = document.getElementById('checkStructure').checked;
            const checkAI = document.getElementById('checkAI').checked;

            if (checkFormat || checkStructure) {
                this.checkResults = FormatChecker.checkDocument(this.parsedDocument);
            } else {
                this.checkResults = {
                    formatIssues: [],
                    contentIssues: [],
                    structureIssues: [],
                    structureInfo: { wordCount: result.value.length, chapters: [] },
                    summary: { errors: 0, warnings: 0, infos: 0, totalIssues: 0 }
                };
            }

            this.displayResults();

            if (checkAI && typeof AIReviewer !== 'undefined') {
                btnLoading.querySelector('.spinner').style.display = 'inline-block';
                btnLoading.innerHTML = '<span class="spinner"></span> AI评审中...';
                
                const aiResult = await AIReviewer.reviewPaper(
                    this.parsedDocument.text,
                    this.checkResults.structureInfo
                );
                
                AIReviewer.displayReviewResult(aiResult);
                this.aiReviewResult = aiResult;
            }

        } catch (error) {
            console.error('解析文件时出错:', error);
            let errorMsg = '解析文件时出错：';
            if (error.message) {
                errorMsg += error.message;
            } else {
                errorMsg += '请确保文件格式正确且未被损坏';
            }
            alert(errorMsg);
        } finally {
            checkBtn.disabled = false;
            btnText.style.display = 'inline';
            btnLoading.style.display = 'none';
        }
    },

    readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                resolve(e.target.result);
            };
            reader.onerror = function(e) {
                reject(new Error('文件读取失败'));
            };
            reader.readAsArrayBuffer(file);
        });
    },

    displayResults() {
        const resultSection = document.getElementById('resultSection');
        resultSection.style.display = 'block';

        const summary = this.checkResults.summary;
        document.getElementById('resultSummary').innerHTML = `
            <div class="summary-item">
                <div class="summary-number" style="color: #ff4757;">${summary.errors}</div>
                <div class="summary-label">错误</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" style="color: #ffa502;">${summary.warnings}</div>
                <div class="summary-label">警告</div>
            </div>
            <div class="summary-item">
                <div class="summary-number" style="color: #3742fa;">${summary.infos}</div>
                <div class="summary-label">提示</div>
            </div>
        `;

        this.renderIssueList('formatIssues', this.checkResults.formatIssues);
        this.renderIssueList('contentIssues', this.checkResults.contentIssues);
        this.renderStructureIssues();
        this.renderStructureInfo();

        resultSection.scrollIntoView({ behavior: 'smooth' });
    },

    renderStructureIssues() {
        const container = document.getElementById('structureInfo');
        const issues = this.checkResults.structureIssues || [];
        const info = this.checkResults.structureInfo;
        
        let html = '<div class="structure-section"><h4>文档结构检测</h4>';
        
        const structureItems = [
            { label: '论文标题', has: info.hasTitle, required: true },
            { label: '中文摘要', has: info.hasChineseAbstract, required: true },
            { label: '英文摘要', has: info.hasEnglishAbstract, required: false },
            { label: '中文关键词', has: info.hasKeywords, required: true },
            { label: '英文关键词', has: info.hasEnglishKeywords, required: false },
            { label: '目录', has: info.hasTOC, required: false },
            { label: '绪论/引言', has: info.hasIntroduction, required: true },
            { label: '结论', has: info.hasConclusion, required: true },
            { label: '参考文献', has: info.hasReferences, required: true },
            { label: '致谢', has: info.hasAcknowledgment, required: false },
            { label: '附录', has: info.hasAppendix, required: false }
        ];
        
        html += '<div class="structure-check-list">';
        structureItems.forEach(item => {
            const statusClass = item.has ? 'status-ok' : (item.required ? 'status-error' : 'status-warning');
            const statusText = item.has ? '✓' : (item.required ? '✗' : '○');
            html += `<div class="structure-check-item ${statusClass}">
                <span class="status-icon">${statusText}</span>
                <span class="status-label">${item.label}</span>
                ${!item.has && item.required ? '<span class="status-tag required">必需</span>' : ''}
                ${!item.has && !item.required ? '<span class="status-tag optional">建议</span>' : ''}
            </div>`;
        });
        html += '</div></div>';
        
        html += '<div class="structure-section"><h4>文档统计</h4>';
        html += '<div class="stats-grid">';
        html += `<div class="stat-item"><span class="stat-value">${info.wordCount}</span><span class="stat-label">总字数</span></div>`;
        html += `<div class="stat-item"><span class="stat-value">${info.paragraphCount}</span><span class="stat-label">段落数</span></div>`;
        html += `<div class="stat-item"><span class="stat-value">${info.chapterCount}</span><span class="stat-label">章节数</span></div>`;
        html += '</div></div>';
        
        if (info.chapters && info.chapters.length > 0) {
            html += '<div class="structure-section"><h4>检测到的章节</h4>';
            html += '<div class="chapter-list">';
            info.chapters.slice(0, 15).forEach((chapter, i) => {
                html += `<div class="chapter-item">
                    <span class="chapter-name">${chapter.name}</span>
                    <span class="chapter-type">${chapter.type === 'numbered' ? '编号章节' : chapter.type === 'special' ? '特殊章节' : '潜在章节'}</span>
                </div>`;
            });
            if (info.chapters.length > 15) {
                html += `<div class="chapter-item more">... 还有 ${info.chapters.length - 15} 个章节</div>`;
            }
            html += '</div></div>';
        }
        
        if (issues.length > 0) {
            html += '<div class="structure-section"><h4>结构问题</h4>';
            html += '<div class="issue-list">';
            issues.forEach(issue => {
                html += `<div class="issue-item ${issue.type}">
                    <div class="issue-header">
                        <span class="issue-title">${issue.title}</span>
                        <span class="issue-badge">${issue.type === 'error' ? '错误' : issue.type === 'warning' ? '警告' : '提示'}</span>
                    </div>
                    <div class="issue-description">${issue.description}</div>
                </div>`;
            });
            html += '</div></div>';
        }
        
        container.innerHTML = html;
    },

    renderIssueList(containerId, issues) {
        const container = document.getElementById(containerId);
        
        if (issues.length === 0) {
            container.innerHTML = `
                <div class="empty-state">
                    <div class="empty-state-icon">✓</div>
                    <p>未发现问题</p>
                </div>
            `;
            return;
        }

        container.innerHTML = issues.map(issue => `
            <div class="issue-item ${issue.type}">
                <div class="issue-header">
                    <span class="issue-title">${issue.title}</span>
                    <span class="issue-badge">${issue.type === 'error' ? '错误' : issue.type === 'warning' ? '警告' : '提示'}</span>
                </div>
                <div class="issue-description">${issue.description}</div>
                <div class="issue-location">位置：${issue.location}</div>
            </div>
        `).join('');
    },

    renderStructureInfo() {
    },

    switchTab(tabName) {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === tabName);
        });
        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.toggle('active', content.id === tabName + 'Tab');
        });
    },

    exportToExcel() {
        if (!this.checkResults) {
            alert('请先运行检测！');
            return;
        }

        const wb = XLSX.utils.book_new();
        const info = this.checkResults.structureInfo;

        const summaryData = [
            ['论文格式检查报告'],
            [''],
            ['文件名', this.currentFile.name],
            ['检测时间', new Date().toLocaleString()],
            [''],
            ['统计摘要'],
            ['错误数', this.checkResults.summary.errors],
            ['警告数', this.checkResults.summary.warnings],
            ['提示数', this.checkResults.summary.infos],
            ['总问题数', this.checkResults.summary.totalIssues],
            [''],
            ['文档信息'],
            ['论文字数', info.wordCount],
            ['段落数量', info.paragraphCount],
            ['章节数量', info.chapterCount],
            [''],
            ['结构检测'],
            ['论文标题', info.hasTitle ? '✓ 已检测到' : '✗ 未检测到'],
            ['中文摘要', info.hasChineseAbstract ? '✓ 已检测到' : '✗ 未检测到'],
            ['英文摘要', info.hasEnglishAbstract ? '✓ 已检测到' : '○ 未检测到'],
            ['中文关键词', info.hasKeywords ? '✓ 已检测到' : '✗ 未检测到'],
            ['目录', info.hasTOC ? '✓ 已检测到' : '○ 未检测到'],
            ['绪论/引言', info.hasIntroduction ? '✓ 已检测到' : '✗ 未检测到'],
            ['结论', info.hasConclusion ? '✓ 已检测到' : '✗ 未检测到'],
            ['参考文献', info.hasReferences ? '✓ 已检测到' : '✗ 未检测到'],
            ['致谢', info.hasAcknowledgment ? '✓ 已检测到' : '○ 未检测到']
        ];
        const ws1 = XLSX.utils.aoa_to_sheet(summaryData);
        XLSX.utils.book_append_sheet(wb, ws1, '检测摘要');

        const structureData = [
            ['序号', '问题类型', '问题描述', '详细说明'],
            ...(this.checkResults.structureIssues || []).map((issue, index) => [
                index + 1,
                issue.type === 'error' ? '错误' : issue.type === 'warning' ? '警告' : '提示',
                issue.title,
                issue.description
            ])
        ];
        const ws2 = XLSX.utils.aoa_to_sheet(structureData);
        XLSX.utils.book_append_sheet(wb, ws2, '结构问题');

        const formatData = [
            ['序号', '问题类型', '问题描述', '详细说明', '位置'],
            ...this.checkResults.formatIssues.map((issue, index) => [
                index + 1,
                issue.type === 'error' ? '错误' : issue.type === 'warning' ? '警告' : '提示',
                issue.title,
                issue.description,
                issue.location
            ])
        ];
        const ws3 = XLSX.utils.aoa_to_sheet(formatData);
        XLSX.utils.book_append_sheet(wb, ws3, '格式问题');

        const contentData = [
            ['序号', '问题类型', '问题描述', '详细说明', '位置'],
            ...this.checkResults.contentIssues.map((issue, index) => [
                index + 1,
                issue.type === 'error' ? '错误' : issue.type === 'warning' ? '警告' : '提示',
                issue.title,
                issue.description,
                issue.location
            ])
        ];
        const ws4 = XLSX.utils.aoa_to_sheet(contentData);
        XLSX.utils.book_append_sheet(wb, ws4, '内容问题');

        if (info.chapters && info.chapters.length > 0) {
            const chapterData = [
                ['序号', '章节名称', '章节类型'],
                ...info.chapters.map((chapter, index) => [
                    index + 1,
                    chapter.name,
                    chapter.type === 'numbered' ? '编号章节' : chapter.type === 'special' ? '特殊章节' : '潜在章节'
                ])
            ];
            const ws5 = XLSX.utils.aoa_to_sheet(chapterData);
            XLSX.utils.book_append_sheet(wb, ws5, '章节列表');
        }

        const fileName = `论文检查报告_${this.currentFile.name.replace('.docx', '')}_${new Date().toLocaleDateString().replace(/\//g, '-')}.xlsx`;
        XLSX.writeFile(wb, fileName);
    }
};

document.addEventListener('DOMContentLoaded', () => {
    PaperChecker.init();
});
