const AIReviewer = {
    config: {
        provider: 'deepseek',
        apiKey: '',
        apiUrl: '',
        model: 'deepseek-chat',
        enabled: true
    },

    providerUrls: {
        openai: 'https://api.openai.com/v1/chat/completions',
        deepseek: 'https://api.deepseek.com/v1/chat/completions',
        zhipu: 'https://open.bigmodel.cn/api/paas/v4/chat/completions',
        qwen: 'https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions',
        moonshot: 'https://api.moonshot.cn/v1/chat/completions'
    },

    providerModels: {
        openai: 'gpt-4o',
        deepseek: 'deepseek-chat',
        zhipu: 'glm-4',
        qwen: 'qwen-max',
        moonshot: 'moonshot-v1-8k'
    },

    init() {
        this.loadConfig();
        this.bindEvents();
        this.updateUI();
    },

    loadConfig() {
        const saved = localStorage.getItem('aiReviewerConfig');
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                this.config = { ...this.config, ...parsed };
            } catch (e) {
                console.error('加载AI配置失败:', e);
            }
        }
    },

    saveConfig() {
        localStorage.setItem('aiReviewerConfig', JSON.stringify(this.config));
    },

    bindEvents() {
        const settingsBtn = document.getElementById('settingsBtn');
        const closeSettings = document.getElementById('closeSettings');
        const settingsModal = document.getElementById('settingsModal');
        const apiProvider = document.getElementById('apiProvider');
        const saveSettings = document.getElementById('saveSettings');
        const testApiBtn = document.getElementById('testApiBtn');

        if (settingsBtn) {
            settingsBtn.addEventListener('click', () => {
                if (settingsModal) {
                    settingsModal.style.display = 'flex';
                    this.updateSettingsForm();
                }
            });
        }

        if (closeSettings) {
            closeSettings.addEventListener('click', () => {
                if (settingsModal) {
                    settingsModal.style.display = 'none';
                }
            });
        }

        if (settingsModal) {
            settingsModal.addEventListener('click', (e) => {
                if (e.target === settingsModal) {
                    settingsModal.style.display = 'none';
                }
            });
        }

        if (apiProvider) {
            apiProvider.addEventListener('change', (e) => {
                const provider = e.target.value;
                const apiUrlGroup = document.getElementById('apiUrlGroup');
                const modelName = document.getElementById('modelName');
                if (apiUrlGroup) {
                    apiUrlGroup.style.display = provider === 'custom' ? 'block' : 'none';
                }
                if (modelName) {
                    modelName.value = this.providerModels[provider] || '';
                }
            });
        }

        if (saveSettings) {
            saveSettings.addEventListener('click', () => {
                const providerEl = document.getElementById('apiProvider');
                const apiKeyEl = document.getElementById('apiKey');
                const apiUrlEl = document.getElementById('apiUrl');
                const modelNameEl = document.getElementById('modelName');
                const enableAIEl = document.getElementById('enableAI');

                if (providerEl) this.config.provider = providerEl.value;
                if (apiKeyEl) this.config.apiKey = apiKeyEl.value;
                if (apiUrlEl && apiUrlEl.value) {
                    this.config.apiUrl = apiUrlEl.value;
                } else {
                    this.config.apiUrl = this.providerUrls[this.config.provider];
                }
                if (modelNameEl) {
                    this.config.model = modelNameEl.value || this.providerModels[this.config.provider];
                }
                if (enableAIEl) this.config.enabled = enableAIEl.checked;
                
                this.saveConfig();
                this.updateUI();
                
                if (settingsModal) {
                    settingsModal.style.display = 'none';
                }
                
                alert('设置已保存！');
            });
        }

        if (testApiBtn) {
            testApiBtn.addEventListener('click', () => this.testConnection());
        }
    },

    updateSettingsForm() {
        const providerEl = document.getElementById('apiProvider');
        const apiKeyEl = document.getElementById('apiKey');
        const apiUrlEl = document.getElementById('apiUrl');
        const modelNameEl = document.getElementById('modelName');
        const enableAIEl = document.getElementById('enableAI');
        const apiUrlGroup = document.getElementById('apiUrlGroup');

        if (providerEl) providerEl.value = this.config.provider;
        if (apiKeyEl) apiKeyEl.value = this.config.apiKey;
        if (apiUrlEl) apiUrlEl.value = this.config.apiUrl;
        if (modelNameEl) modelNameEl.value = this.config.model || this.providerModels[this.config.provider];
        if (enableAIEl) enableAIEl.checked = this.config.enabled;
        if (apiUrlGroup) apiUrlGroup.style.display = this.config.provider === 'custom' ? 'block' : 'none';
    },

    updateUI() {
        const aiCheckbox = document.getElementById('checkAI');
        if (aiCheckbox) {
            aiCheckbox.disabled = !this.config.enabled || !this.config.apiKey;
        }
    },

    async testConnection() {
        const statusEl = document.getElementById('aiStatus');
        if (!statusEl) return;
        
        statusEl.innerHTML = '<span class="status-testing">正在测试连接...</span>';

        try {
            const providerEl = document.getElementById('apiProvider');
            const apiKeyEl = document.getElementById('apiKey');
            const modelNameEl = document.getElementById('modelName');
            const apiUrlEl = document.getElementById('apiUrl');

            const provider = providerEl ? providerEl.value : 'deepseek';
            const apiKey = apiKeyEl ? apiKeyEl.value : '';
            const model = modelNameEl ? (modelNameEl.value || this.providerModels[provider]) : this.providerModels[provider];
            const url = provider === 'custom' 
                ? (apiUrlEl ? apiUrlEl.value : '')
                : this.providerUrls[provider];

            if (!apiKey) {
                throw new Error('请输入API Key');
            }

            if (!url) {
                throw new Error('请输入API地址');
            }

            const response = await this.callAPI(url, apiKey, model, [
                { role: 'user', content: '你好，请回复"连接成功"' }
            ], 50);

            if (response) {
                statusEl.innerHTML = '<span class="status-success">✓ 连接成功！</span>';
            }
        } catch (error) {
            statusEl.innerHTML = '<span class="status-error">✗ 连接失败: ' + error.message + '</span>';
        }
    },

    async callAPI(url, apiKey, model, messages, maxTokens) {
        maxTokens = maxTokens || 4000;
        
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + apiKey
        };

        const body = {
            model: model,
            messages: messages,
            max_tokens: maxTokens,
            temperature: 0.7
        };

        const response = await fetch(url, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(body)
        });

        if (!response.ok) {
            let errorMsg = 'HTTP ' + response.status;
            try {
                const errorData = await response.json();
                if (errorData.error && errorData.error.message) {
                    errorMsg = errorData.error.message;
                }
            } catch (e) {
                // ignore
            }
            throw new Error(errorMsg);
        }

        const data = await response.json();
        
        if (data.choices && data.choices[0] && data.choices[0].message) {
            return data.choices[0].message.content;
        }
        
        return '';
    },

    buildReviewPrompt(paperContent, structureInfo) {
        const wordCount = structureInfo.wordCount || 0;
        const paragraphCount = structureInfo.paragraphCount || 0;
        const chapterCount = structureInfo.chapterCount || 0;
        const chapters = structureInfo.chapters || [];
        const chapterNames = chapters.map(function(c) { return c.name; }).join('、') || '未检测到';

        return '你是一位经验丰富的大学毕业论文指导老师，请对以下本科毕业论文进行全面评审。\n\n' +
            '## 评审要求\n' +
            '请严格按照《本科毕业论文撰写规范》进行评审，从以下几个方面进行详细分析：\n\n' +
            '### 1. 格式规范性评审\n' +
            '- 标题格式：论文标题、章节标题是否规范\n' +
            '- 摘要格式：中英文摘要是否符合要求\n' +
            '- 关键词格式：关键词数量和格式是否规范\n' +
            '- 正文格式：段落缩进、字体字号、行距等\n' +
            '- 参考文献格式：是否符合GB/T 7714标准\n\n' +
            '### 2. 论文结构评审\n' +
            '- 整体结构是否完整（摘要、目录、绪论、正文、结论、参考文献、致谢）\n' +
            '- 章节安排是否合理\n' +
            '- 各章节篇幅比例是否恰当\n\n' +
            '### 3. 内容质量评审\n' +
            '- 研究问题是否明确\n' +
            '- 研究方法是否科学合理\n' +
            '- 论证过程是否严谨\n' +
            '- 数据分析是否充分\n' +
            '- 结论是否合理\n\n' +
            '### 4. 逻辑性评审\n' +
            '- 全文逻辑是否通顺\n' +
            '- 章节之间衔接是否自然\n' +
            '- 论点论据是否一致\n\n' +
            '### 5. 学术规范评审\n' +
            '- 引用是否规范\n' +
            '- 是否存在抄袭嫌疑\n' +
            '- 专业术语使用是否准确\n\n' +
            '## 论文基本信息\n' +
            '- 字数：' + wordCount + '字\n' +
            '- 段落数：' + paragraphCount + '段\n' +
            '- 章节数：' + chapterCount + '章\n' +
            '- 检测到的章节：' + chapterNames + '\n\n' +
            '## 论文内容\n' +
            paperContent.substring(0, 12000) + '\n\n' +
            '## 输出要求\n' +
            '请以JSON格式输出评审结果，格式如下：\n' +
            '{\n' +
            '    "overallScore": 85,\n' +
            '    "overallComment": "总体评价...",\n' +
            '    "categories": [\n' +
            '        {\n' +
            '            "name": "格式规范性",\n' +
            '            "score": 80,\n' +
            '            "issues": [\n' +
            '                {"type": "error", "title": "问题标题", "description": "问题描述", "suggestion": "修改建议"}\n' +
            '            ],\n' +
            '            "comment": "分类评价..."\n' +
            '        }\n' +
            '    ],\n' +
            '    "suggestions": ["改进建议1", "改进建议2"],\n' +
            '    "highlights": ["优点1", "优点2"]\n' +
            '}\n\n' +
            '请确保输出的是有效的JSON格式，不要包含其他内容。';
    },

    async reviewPaper(paperContent, structureInfo) {
        if (!this.config.enabled || !this.config.apiKey) {
            return {
                success: false,
                error: 'AI评审未启用或未配置API Key，请点击右上角设置按钮进行配置',
                data: null
            };
        }

        try {
            const url = this.config.apiUrl || this.providerUrls[this.config.provider];
            const prompt = this.buildReviewPrompt(paperContent, structureInfo);
            
            const response = await this.callAPI(
                url,
                this.config.apiKey,
                this.config.model,
                [{ role: 'user', content: prompt }],
                4000
            );

            let reviewData;
            try {
                const jsonMatch = response.match(/\{[\s\S]*\}/);
                if (jsonMatch) {
                    reviewData = JSON.parse(jsonMatch[0]);
                } else {
                    throw new Error('无法解析AI响应');
                }
            } catch (e) {
                reviewData = {
                    overallScore: 0,
                    overallComment: response,
                    categories: [],
                    suggestions: [],
                    highlights: [],
                    rawResponse: response
                };
            }

            return {
                success: true,
                data: reviewData
            };
        } catch (error) {
            return {
                success: false,
                error: error.message,
                data: null
            };
        }
    },

    displayReviewResult(result) {
        const container = document.getElementById('aiReview');
        if (!container) return;
        
        if (!result.success) {
            container.innerHTML = '<div class="ai-error">' +
                '<div class="error-icon">⚠️</div>' +
                '<h4>AI评审失败</h4>' +
                '<p>' + result.error + '</p>' +
                '<p class="error-hint">请检查API配置是否正确，或稍后重试</p>' +
                '</div>';
            return;
        }

        const data = result.data;
        let html = '';

        if (data.overallScore) {
            let scoreClass = 'score-medium';
            if (data.overallScore >= 80) scoreClass = 'score-good';
            else if (data.overallScore < 60) scoreClass = 'score-poor';
            
            html += '<div class="ai-overall">' +
                '<div class="overall-score ' + scoreClass + '">' +
                '<span class="score-number">' + data.overallScore + '</span>' +
                '<span class="score-label">综合评分</span>' +
                '</div>' +
                '<div class="overall-comment">' +
                '<h4>总体评价</h4>' +
                '<p>' + (data.overallComment || '暂无评价') + '</p>' +
                '</div>' +
                '</div>';
        }

        if (data.highlights && data.highlights.length > 0) {
            html += '<div class="ai-section highlights">' +
                '<h4>👍 论文优点</h4>' +
                '<ul>';
            for (let i = 0; i < data.highlights.length; i++) {
                html += '<li>' + data.highlights[i] + '</li>';
            }
            html += '</ul></div>';
        }

        if (data.categories && data.categories.length > 0) {
            html += '<div class="ai-categories">';
            for (let i = 0; i < data.categories.length; i++) {
                const category = data.categories[i];
                let catScoreClass = 'score-medium';
                if (category.score >= 80) catScoreClass = 'score-good';
                else if (category.score < 60) catScoreClass = 'score-poor';
                
                html += '<div class="ai-category">' +
                    '<div class="category-header">' +
                    '<h5>' + category.name + '</h5>' +
                    '<span class="category-score ' + catScoreClass + '">' + category.score + '分</span>' +
                    '</div>';
                
                if (category.comment) {
                    html += '<p class="category-comment">' + category.comment + '</p>';
                }
                
                if (category.issues && category.issues.length > 0) {
                    html += '<div class="category-issues">';
                    for (let j = 0; j < category.issues.length; j++) {
                        const issue = category.issues[j];
                        const issueIcon = issue.type === 'error' ? '❌' : (issue.type === 'warning' ? '⚠️' : 'ℹ️');
                        html += '<div class="ai-issue ' + issue.type + '">' +
                            '<span class="issue-type">' + issueIcon + '</span>' +
                            '<div class="issue-content">' +
                            '<strong>' + issue.title + '</strong>' +
                            '<p>' + issue.description + '</p>';
                        if (issue.suggestion) {
                            html += '<p class="suggestion">💡 建议：' + issue.suggestion + '</p>';
                        }
                        html += '</div></div>';
                    }
                    html += '</div>';
                }
                html += '</div>';
            }
            html += '</div>';
        }

        if (data.suggestions && data.suggestions.length > 0) {
            html += '<div class="ai-section suggestions">' +
                '<h4>📝 改进建议</h4>' +
                '<ol>';
            for (let i = 0; i < data.suggestions.length; i++) {
                html += '<li>' + data.suggestions[i] + '</li>';
            }
            html += '</ol></div>';
        }

        if (data.rawResponse) {
            html += '<div class="ai-section raw-response">' +
                '<h4>📄 AI原始响应</h4>' +
                '<div class="raw-content">' + data.rawResponse + '</div>' +
                '</div>';
        }

        container.innerHTML = html || '<div class="ai-empty">暂无评审结果</div>';
    }
};

document.addEventListener('DOMContentLoaded', function() {
    AIReviewer.init();
});
