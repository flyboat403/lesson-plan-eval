const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, 
        AlignmentType, BorderStyle, WidthType, ShadingType, 
        HeadingLevel, LevelFormat, Header, Footer, PageNumber,
        VerticalAlign } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const headerB = { top: { style: BorderStyle.SINGLE, size: 2, color: "2E75B6" }, bottom: { style: BorderStyle.SINGLE, size: 2, color: "2E75B6" }, left: border, right: border };

function cell(text, w, hdr, align) {
    return new TableCell({
        borders: hdr ? headerB : borders,
        width: { size: w, type: WidthType.DXA },
        shading: hdr ? { fill: "D5E8F0", type: ShadingType.CLEAR } : undefined,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ alignment: align || AlignmentType.LEFT, children: [new TextRun({ text, bold: !!hdr, font: "Arial", size: hdr ? 22 : 20 })] })]
    });
}

function tRow(cells, bg) {
    return new TableRow({ children: cells.map(c => cell(c.text, c.w, c.hdr, c.align)), 
        tableHeader: true, shading: bg ? { fill: bg, type: ShadingType.CLEAR } : undefined });
}

const doc = new Document({
    styles: {
        default: { document: { run: { font: "Arial", size: 22 } } },
        paragraphStyles: [
            { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 36, bold: true, font: "Arial", color: "2E75B6" }, paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
            { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: "Arial", color: "2E75B6" }, paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
            { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, font: "Arial" }, paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 2 } },
        ]
    },
    numbering: {
        config: [
            { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
            { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }
        ]
    },
    sections: [{
        properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
        headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "教案审计报告 | AI教研员评估系统", font: "Arial", size: 18, color: "666666" })] })] }) },
        footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "第 ", font: "Arial", size: 18 }), new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18 }), new TextRun({ text: " 页", font: "Arial", size: 18 })] })] }) },
        children: [
            new Paragraph({ heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "教案审计报告", bold: true, font: "Arial", size: 44, color: "2E75B6" })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 }, children: [new TextRun({ text: "心理健康与职业生涯 — 三省吾身 心知其意", font: "Arial", size: 24 })] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("评级概览")] }),
            new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2340, 7020], rows: [
                tRow([{ text: "综合评级", w: 2340, hdr: true }, { text: "⭐⭐⭐☆☆ 合格", w: 7020 }]),
                tRow([{ text: "总分", w: 2340, hdr: true }, { text: "66/100分", w: 7020 }]),
                tRow([{ text: "总体结论", w: 2340, hdr: true }, { text: "合格 — 存在明显问题，需重点修改后使用", w: 7020 }]),
                tRow([{ text: "课程类型", w: 2340, hdr: true }, { text: "公共基础课（不检查工作过程标注）", w: 7020 }]),
                tRow([{ text: "诊断模式", w: 2340, hdr: true }, { text: "全面逆向审计", w: 7020 }])
            ]}),

            new Paragraph({ spacing: { before: 300 }, children: [] }),
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("分项得分表")] }),
            new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 1200, 1200, 5160], rows: [
                tRow([{ text: "维度", w: 2000, hdr: true, align: AlignmentType.CENTER }, { text: "得分", w: 1200, hdr: true, align: AlignmentType.CENTER }, { text: "满分", w: 1200, hdr: true, align: AlignmentType.CENTER }, { text: "主要问题", w: 5160, hdr: true }]),
                tRow([{ text: "守门检查", w: 2000 }, { text: "13", w: 1200, align: AlignmentType.CENTER }, { text: "15", w: 1200, align: AlignmentType.CENTER }, { text: "课堂类型标注不规范", w: 5160 }]),
                tRow([{ text: "目标审计", w: 2000 }, { text: "15", w: 1200, align: AlignmentType.CENTER }, { text: "20", w: 1200, align: AlignmentType.CENTER }, { text: "核心问题偏离、目标模糊", w: 5160 }]),
                tRow([{ text: "问题链诊断", w: 2000 }, { text: "12", w: 1200, align: AlignmentType.CENTER }, { text: "15", w: 1200, align: AlignmentType.CENTER }, { text: "问题递进链不完整", w: 5160 }]),
                tRow([{ text: "活动反假", w: 2000 }, { text: "16", w: 1200, align: AlignmentType.CENTER }, { text: "20", w: 1200, align: AlignmentType.CENTER }, { text: "考核标准不够明确", w: 5160 }]),
                tRow([{ text: "辅助审计", w: 2000 }, { text: "11", w: 1200, align: AlignmentType.CENTER }, { text: "20", w: 1200, align: AlignmentType.CENTER }, { text: "讲授过长、缺板书、反思空置", w: 5160 }]),
                tRow([{ text: "学情审计", w: 2000 }, { text: "9", w: 1200, align: AlignmentType.CENTER }, { text: "10", w: 1200, align: AlignmentType.CENTER }, { text: "差异应对策略不足", w: 5160 }]),
                tRow([{ text: "总计", w: 2000, hdr: true, align: AlignmentType.CENTER }, { text: "66", w: 1200, hdr: true, align: AlignmentType.CENTER }, { text: "100", w: 1200, hdr: true, align: AlignmentType.CENTER }, { text: "合格（60-74分）", w: 5160, hdr: true }], "E6F3E6")
            ]}),

            new Paragraph({ spacing: { before: 300 }, children: [] }),
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("值得保留的亮点")] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "学情分析七维度全覆盖：", bold: true }), new TextRun(" 从来源、学业基础、认知能力、学习特点、专业特性、职业面向、发展诉求七维度分析，辅以数据。")] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "活动设计富有创新性：", bold: true }), new TextRun(" AI画像反向验证、自我探索卡片传递多轮迭代。")] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "课后拓展设计出色：", bold: true }), new TextRun(" 志愿服务中的遇见自己，实现知行合一。")] }),

            new Paragraph({ spacing: { before: 300 }, children: [] }),
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("核心硬伤与改进建议")] }),
            new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [500, 1300, 1200, 700, 2600, 2760], rows: [
                tRow([{ text: "#", w: 500, hdr: true, align: AlignmentType.CENTER }, { text: "问题位置", w: 1300, hdr: true }, { text: "类型", w: 1200, hdr: true }, { text: "扣分", w: 700, hdr: true, align: AlignmentType.CENTER }, { text: "诊断分析", w: 2600, hdr: true }, { text: "改进建议", w: 2760, hdr: true }]),
                tRow([{ text: "1", w: 500, align: AlignmentType.CENTER }, { text: "总议题", w: 1300 }, { text: "核心问题偏离", w: 1200 }, { text: "-5", w: 700, align: AlignmentType.CENTER }, { text: "人人都能成为社会有用人才过于宏大", w: 2600 }, { text: "改为：如何全面客观地认识自己？", w: 2760 }]),
                tRow([{ text: "2", w: 500, align: AlignmentType.CENTER }, { text: "学科目标", w: 1300 }, { text: "目标模糊", w: 1200 }, { text: "-3", w: 700, align: AlignmentType.CENTER }, { text: "掌握分析信息能力无法观测", w: 2600 }, { text: "改为能运用自我觉察法评价", w: 2760 }]),
                tRow([{ text: "3", w: 500, align: AlignmentType.CENTER }, { text: "思政目标", w: 1300 }, { text: "思政融入生硬", w: 1200 }, { text: "-2", w: 700, align: AlignmentType.CENTER }, { text: "脱离本课主题生硬贴标签", w: 2600 }, { text: "结合护理专业尊重患者差异", w: 2760 }]),
                tRow([{ text: "4", w: 500, align: AlignmentType.CENTER }, { text: "析心理环节", w: 1300 }, { text: "讲授过长", w: 1200 }, { text: "-4", w: 700, align: AlignmentType.CENTER }, { text: "35分钟含大量讲解违反注意力曲线", w: 2600 }, { text: "拆分为讲8分+活动27分", w: 2760 }]),
                tRow([{ text: "5", w: 500, align: AlignmentType.CENTER }, { text: "导心知提问", w: 1300 }, { text: "伪互动风险", w: 1200 }, { text: "-2", w: 700, align: AlignmentType.CENTER }, { text: "问题过于抽象缺乏回答支架", w: 2600 }, { text: "改为哪吒眼中自己与他人不同", w: 2760 }]),
                tRow([{ text: "6", w: 500, align: AlignmentType.CENTER }, { text: "板书设计", w: 1300 }, { text: "缺失主板书", w: 1200 }, { text: "-3", w: 700, align: AlignmentType.CENTER }, { text: "未设计主板书缺知识锚点", w: 2600 }, { text: "设计三维框架+三方法流程图", w: 2760 }]),
                tRow([{ text: "7", w: 500, align: AlignmentType.CENTER }, { text: "教学反思", w: 1300 }, { text: "缺反思框架", w: 1200 }, { text: "-2", w: 700, align: AlignmentType.CENTER }, { text: "教学反思表格完全空置", w: 2600 }, { text: "增加成功/问题/措施框架", w: 2760 }])
            ]}),

            new Paragraph({ spacing: { before: 300 }, children: [] }),
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("专家重构：问题链优化")] }),
            new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1800, 1400, 2000, 1600, 1600], rows: [
                tRow([{ text: "原问题", w: 1800, hdr: true }, { text: "认知层级/诊断", w: 1400, hdr: true }, { text: "优化后的问题", w: 2000, hdr: true }, { text: "可追加追问", w: 1600, hdr: true }, { text: "训练能力", w: 1600, hdr: true }]),
                tRow([{ text: "我和他人眼中的我是什么关系", w: 1800 }, { text: "分析/过于宏大", w: 1400 }, { text: "哪吒眼中自己与他人有何不同", w: 2000 }, { text: "这种不同带来什么困扰", w: 1600 }, { text: "观察与对比能力", w: 1600 }]),
                tRow([{ text: "探秘我是谁如何发现独特自我", w: 1800 }, { text: "探究/缺少脚手架", w: 1400 }, { text: "用三个词描述你眼中的自己", w: 2000 }, { text: "对应生理心理还是社会层面", w: 1600 }, { text: "自我觉察能力", w: 1600 }]),
                tRow([{ text: "多棱镜中的我如何借助方法看清自己", w: 1800 }, { text: "方法层/与活动脱节", w: 1400 }, { text: "你发现了哪些没注意到的特点", w: 2000 }, { text: "这个发现让认识有什么改变", w: 1600 }, { text: "反思与整合能力", w: 1600 }]),
            ]}),

            new Paragraph({ spacing: { before: 300 }, children: [] }),
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("优化后的课堂结构简表")] }),
            new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1200, 2400, 2400, 2160, 1200], rows: [
                tRow([{ text: "环节", w: 1200, hdr: true, align: AlignmentType.CENTER }, { text: "教师活动", w: 2400, hdr: true }, { text: "学生活动", w: 2400, hdr: true }, { text: "设计意图", w: 2160, hdr: true }, { text: "时间", w: 1200, hdr: true, align: AlignmentType.CENTER }]),
                tRow([{ text: "导心知", w: 1200 }, { text: "播放哪吒片段引导分析自我认知与他人评价的冲突", w: 2400 }, { text: "观看视频思考分享关键词", w: 2400 }, { text: "创设情境激发兴趣", w: 2160 }, { text: "5 min", w: 1200, align: AlignmentType.CENTER }]),
                tRow([{ text: "析心理-讲", w: 1200 }, { text: "微讲授三个维度含义与示例", w: 2400 }, { text: "听讲记录结合理解", w: 2400 }, { text: "建立认知框架", w: 2160 }, { text: "8 min", w: 1200, align: AlignmentType.CENTER }]),
                tRow([{ text: "析心理-做1", w: 1200 }, { text: "组织我是谁活动自我描述同伴修改AI验证", w: 2400 }, { text: "书写三维度描述卡同伴互评观察AI画像", w: 2400 }, { text: "体验三维度认识自我", w: 2160 }, { text: "15 min", w: 1200, align: AlignmentType.CENTER }]),
                tRow([{ text: "析心理-做2", w: 1200 }, { text: "组织自我探索活动评价传递卡片畅想十年", w: 2400 }, { text: "卡片传递互评闭眼畅想记录分享", w: 2400 }, { text: "体验三种认识方法找到差距", w: 2160 }, { text: "12 min", w: 1200, align: AlignmentType.CENTER }]),
                tRow([{ text: "评心效", w: 1200 }, { text: "知识回顾希沃游戏测评价值升华", w: 2400 }, { text: "回答参与游戏情感共鸣", w: 2400 }, { text: "检验掌握升华主题", w: 2160 }, { text: "5 min", w: 1200, align: AlignmentType.CENTER }]),
            ]}),

            new Paragraph({ spacing: { before: 300 }, children: [] }),
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("教学改进路线图")] }),
            new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: "微观修正", color: "28A745" }), new TextRun("（本课直接使用）")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun(" 总议题改为核心问题：如何全面、客观地认识自己？")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun(" 析心理拆分为讲8分+活动27分")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun(" 增加主板书设计：三维框架+三方法流程图")] }),
            new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200 }, children: [new TextRun({ text: "中观重构", color: "FFC107" }), new TextRun("（本单元优化）")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun(" 重写目标使用精确行为动词")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun(" 修正思政融入方式结合护理专业特性")] }),
            new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 200 }, children: [new TextRun({ text: "宏观研讨", color: "DC3545" }), new TextRun("（建议教研组讨论）")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun(" 公共基础课如何体现岗课对接？")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun(" 1课时是否足够？是否拆分为2课时？")] }),

            new Paragraph({ spacing: { before: 400 }, children: [] }),
            new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "CCCCCC", space: 1 } }, children: [] }),
            new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 200 }, children: [new TextRun({ text: "生成时间: 2026-05-11 | 审计模型: Vocational Lesson Audit Agent v0.2.0", font: "Arial", size: 18, color: "666666", italics: true })] }),
        ]
    }]
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync("教案审计报告_发现自我.docx", buffer);
    console.log("✅ 报告生成成功：教案审计报告_发现自我.docx");
});