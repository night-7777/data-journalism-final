// Fill in 表格.doc template with project data
const docx = require('docx');
const fs = require('fs');
const path = require('path');

const {
  Document, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle,
  PageBreak, ImageRun, Packer
} = docx;

// Colors matching the original template (green theme)
const DARK_GREEN = '1B5E20';
const LIGHT_GREEN = '4CAF50';
const HEADER_BG = '2E7D32';
const ZEBRA = 'F1F8E9';

// Helper: standard cell
function makeCell(text, opts = {}) {
  const { bold = false, color = '000000', bg = null, align = AlignmentType.LEFT, size = 22, width = null, font = '宋体' } = opts;
  return new TableCell({
    width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
    shading: bg ? { fill: bg } : undefined,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({
        text: text || '',
        bold, color, size, font: { eastAsia: font, ascii: font, hAnsi: font }
      })]
    })]
  });
}

function makeRichCell(paragraphs, opts = {}) {
  const { bg = null, width = null } = opts;
  return new TableCell({
    width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
    shading: bg ? { fill: bg } : undefined,
    children: paragraphs
  });
}

function p(text, opts = {}) {
  const { bold = false, color = '000000', size = 22, align = AlignmentType.LEFT, font = '宋体' } = opts;
  return new Paragraph({
    spacing: { before: 60, after: 60, line: 320 },
    alignment: align,
    children: [new TextRun({
      text,
      bold, color, size,
      font: { eastAsia: font, ascii: font, hAnsi: font }
    })]
  });
}

function emptyP() { return new Paragraph({ children: [new TextRun('')] }); }

// Members list
const members = [
  { id: 'P241012401', name: '甘校松' },
  { id: 'P241012428', name: '安欣' },
  { id: 'P241012460', name: '邓继诚' },
  { id: 'P241012407', name: '方子轩' }
];

// Project URL
const PROJECT_URL = 'https://604ad7e966da45e7b9927bbd5de18977.app.codebuddy.work';

// ====== BUILD TABLES ======

// Header table: 学校 / 学院 / 作品说明书
const headerTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  alignment: AlignmentType.CENTER,
  borders: {
    top: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
    bottom: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
    left: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
    right: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
    insideHorizontal: { style: BorderStyle.NONE, size: 0 },
    insideVertical: { style: BorderStyle.NONE, size: 0 }
  },
  rows: [new TableRow({
    children: [
      new TableCell({
        width: { size: 33, type: WidthType.PERCENTAGE },
        children: [p('西北民族大学', { bold: true, size: 26, align: AlignmentType.CENTER, color: DARK_GREEN })]
      }),
      new TableCell({
        width: { size: 34, type: WidthType.PERCENTAGE },
        children: [p('新闻传播学院', { bold: true, size: 26, align: AlignmentType.CENTER, color: DARK_GREEN })]
      }),
      new TableCell({
        width: { size: 33, type: WidthType.PERCENTAGE },
        children: [p('作品说明书', { bold: true, size: 26, align: AlignmentType.CENTER, color: DARK_GREEN })]
      })
    ]
  })]
});

// Project info table: 学期 / 课程名称 / 专业 / 班级 / 作品名称
const infoTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
    bottom: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
    left: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
    right: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
    insideVertical: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' }
  },
  rows: [
    ['学期', '2025-2026-2', '课程名称', '数据新闻理论与实践'],
    ['专业', '新闻学', '班级', '2024级新闻学3班'],
    ['作品名称', { text: '中国森林、湿地覆盖率各省对比', span: 3 }]
  ].map((row, i) => new TableRow({
    children: row.map((cell, j) => {
      if (typeof cell === 'string') {
        const isLabel = j % 2 === 0;
        return makeCell(cell, {
          bold: isLabel, bg: isLabel ? ZEBRA : null, align: AlignmentType.CENTER,
          color: isLabel ? DARK_GREEN : '000000', width: 15
        });
      } else {
        return new TableCell({
          columnSpan: cell.span,
          width: { size: 70, type: WidthType.PERCENTAGE },
          children: [p(cell.text, { bold: true, size: 24, align: AlignmentType.CENTER, color: DARK_GREEN })]
        });
      }
    })
  }))
});

// Members table: 成员名单 / 学号 / 姓名
const membersTable = new Table({
  width: { size: 100, type: WidthType.PERCENTAGE },
  borders: {
    top: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
    bottom: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
    left: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
    right: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
    insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
    insideVertical: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' }
  },
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        makeCell('成员名单', { bold: true, bg: HEADER_BG, color: 'FFFFFF', align: AlignmentType.CENTER, size: 24, width: 25 }),
        makeCell('学号', { bold: true, bg: HEADER_BG, color: 'FFFFFF', align: AlignmentType.CENTER, size: 24, width: 25 }),
        makeCell('姓名', { bold: true, bg: HEADER_BG, color: 'FFFFFF', align: AlignmentType.CENTER, size: 24, width: 25 }),
        makeCell('任务分工', { bold: true, bg: HEADER_BG, color: 'FFFFFF', align: AlignmentType.CENTER, size: 24, width: 25 })
      ]
    }),
    ...members.map((m, i) => new TableRow({
      children: [
        makeCell(`成员${i+1}`, { align: AlignmentType.CENTER, bg: ZEBRA }),
        makeCell(m.id, { align: AlignmentType.CENTER }),
        makeCell(m.name, { align: AlignmentType.CENTER, bold: true }),
        makeCell(['数据检索与整理', '图表配色与交互', '作品说明撰写、统筹', '数据校验与查漏补缺'][i], { align: AlignmentType.LEFT, size: 20 })
      ]
    }))
  ]
});

// Section content helper
function sectionBlock(title, content) {
  return [
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
        bottom: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
        left: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
        right: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
        insideHorizontal: { style: BorderStyle.NONE, size: 0 },
        insideVertical: { style: BorderStyle.NONE, size: 0 }
      },
      rows: [new TableRow({
        children: [new TableCell({
          width: { size: 100, type: WidthType.PERCENTAGE },
          shading: { fill: HEADER_BG },
          children: [p(title, { bold: true, size: 28, color: 'FFFFFF', align: AlignmentType.CENTER })]
        })]
      })]
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
        bottom: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
        left: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
        right: { style: BorderStyle.SINGLE, size: 8, color: '999999' },
        insideHorizontal: { style: BorderStyle.NONE, size: 0 },
        insideVertical: { style: BorderStyle.NONE, size: 0 }
      },
      rows: [new TableRow({
        children: [new TableCell({
          width: { size: 100, type: WidthType.PERCENTAGE },
          children: content
        })]
      })]
    })
  ];
}

// Content sections
const workIntro = sectionBlock('一、作品简介', [
  p('《中国森林、湿地覆盖率各省对比》是一组聚焦中国 31 个省级行政区生态资源空间分布格局的数据新闻作品，涵盖一个交互式信息可视化网页与配套图文报告。', { size: 22 }),
  p('选题背景：森林与湿地是衡量区域生态禀赋与生物多样性的两大核心指标。当前公众讨论多停留在「全国均值」，缺乏对省级差异、东南—西北分异规律的直观感知。本作品试图用统一的数据底板、可交互的图表与可对比的视角，把「绿色中国」的内部差异讲清楚。', { size: 22 }),
  p('选题意义：① 为生态文明建设成果提供省域尺度的可视化呈现；② 通过森林与湿地的耦合分析，揭示两类生态资源的「此消彼长」与「共生共存」关系；③ 沉淀一套可复用的数据新闻教学范式（选题—采集—清洗—可视化—解读—传播）。', { size: 22 })
]);

const workLink = sectionBlock('二、作品链接', [
  p('（一）可在线访问的交互式信息可视化网页：', { bold: true, size: 22, color: DARK_GREEN }),
  p(`链接：${PROJECT_URL}`, { size: 22 }),
  p('本作品为半动态交互式网页，共设 4 个 Tab：总览、排名对比、省市检索、双省对比，覆盖数字动画、雷达图、柱状图、散点图、检索表、双省对比等多种可视化形态。', { size: 22 })
]);

const workNotes = sectionBlock('三、创作手记', [
  p('（一）内容结构', { bold: true, size: 24, color: DARK_GREEN }),
  p('作品采用「引言—指标解释—全国概览—分省对比—区域分析—双省深度对比—结论」的递进结构。', { size: 22 }),
  p('（二）数据来源', { bold: true, size: 24, color: DARK_GREEN }),
  p('1. 森林覆盖率：第九次全国森林资源清查（截至 2024 年）公开数据。', { size: 22 }),
  p('2. 湿地面积：国家统计局《第一次全国湿地资源调查（1995-2003）》汇总数据，单位：千公顷（不含水稻田湿地）。', { size: 22 }),
  p('3. 省辖区面积：自然资源部公开地理数据。', { size: 22 }),
  p('（三）数据挖掘', { bold: true, size: 24, color: DARK_GREEN }),
  p('基于 31 省原始数据计算了「湿地覆盖率 = 湿地面积 ÷ 省辖区面积 × 100%」指标；按七大地理分区（华东/华南/西南/华中/华北/东北/西北）聚合分析；构建森林覆盖率 × 湿地覆盖率二维分类模型，形成「双高/林高湿低/湿高林低/双低」四象限。', { size: 22 }),
  p('（四）数据分析', { bold: true, size: 24, color: DARK_GREEN }),
  p('关键发现：① 福建（66.80%）、江西、广西、浙江构成「东南林区」第一梯队；② 上海（50.74%）、天津（14.44%）位列湿地覆盖率前两位，主因滨海滩涂与河口湿地；③ 西藏、青海、黑龙江为湿地面积总量三强，但单位面积湿地率仍受地形约束。', { size: 22 }),
  p('（五）数据可视化', { bold: true, size: 24, color: DARK_GREEN }),
  p('采用 ECharts 5 + HTML/CSS3 技术栈。配色以森林绿、湿地蓝为主色，辅以金色与暖橙强调关键数据；所有图表均支持 hover 高亮、点击钻取、滑块缩放、brush 框选、雷达对比、动画数字计数器等多种交互形式。', { size: 22 }),
  p('（六）心路历程', { bold: true, size: 24, color: DARK_GREEN }),
  p('选题阶段曾考虑加入「草原覆盖率」与「物种多样性指数」两项指标，后因数据可得性与可视化复杂度权衡放弃。数据采集阶段遭遇部分省份湿地数据缺失，最终选用国家林业局「第一次全国湿地调查」作为权威统一来源。可视化阶段尝试过地图热力图，因加载速度与配色平衡问题调整为二维散点 + 区域进度条方案。', { size: 22 })
]);

const toolDesc = sectionBlock('四、使用工具及使用阐释', [
  p('1. 数据检索：百度搜索、必应搜索、WebFetch 抓取公开统计报告。', { size: 22 }),
  p('2. 数据整理：Excel（原始数据清洗 + 公式计算湿地覆盖率）。', { size: 22 }),
  p('3. 可视化开发：ECharts 5（图表）+ HTML5/CSS3（页面框架与样式）+ JavaScript（交互逻辑）。', { size: 22 }),
  p('4. 文档撰写：Node.js + docx 库（Word 报告生成）、FPDF2（PDF 报告生成）、Python 3.13。', { size: 22 }),
  p('5. 部署与传播：CloudStudio 静态站点托管（在线访问）、GitHub Pages（开源版本）。', { size: 22 })
]);

const materialDesc = sectionBlock('五、素材使用情况说明', [
  p('1. 图表数据：均来自第九次全国森林资源清查、国家统计局湿地调查；地理区划采用国家统计局七大分区标准。', { size: 22 }),
  p('2. 配色方案：参考国家林草局官方色卡与 UNEP「绿色中国」生态报告视觉语言。', { size: 22 }),
  p('3. 字体：宋体（正文）、微软雅黑（标题）、Arial Unicode MS（西文与数字）。', { size: 22 }),
  p('4. 互动组件：ECharts 图表库（Apache 2.0 协议），免费用于非商业用途。', { size: 22 }),
  p('5. 全部素材为公开数据或开源库，无版权争议。', { size: 22 })
]);

const highlights = sectionBlock('六、作品亮点和不足之处', [
  p('（一）亮点', { bold: true, size: 24, color: DARK_GREEN }),
  p('1. 双指标耦合：首次将「森林覆盖率」与「湿地覆盖率」置于同一坐标系进行 31 省比较，形成四象限分类模型；', { size: 22 }),
  p('2. 半动态交互：4 个 Tab 涵盖 6+ 种图表形态，支持搜索、排序、框选、双省对比等十余种交互动作；', { size: 22 }),
  p('3. 数据透明：所有数据来源、计算公式、更新时间均在作品中明示，符合数据新闻「可验证性」原则；', { size: 22 }),
  p('4. 视觉规范：统一的绿/蓝主色调 + 四大地理分区的差异化强调色，提升可读性与记忆点。', { size: 22 }),
  p('（二）不足之处', { bold: true, size: 24, color: DARK_GREEN }),
  p('1. 数据时效性：湿地数据仍以 2003 年首次调查为底，2010 年后的二调数据未在全国层面完整公开；', { size: 22 }),
  p('2. 移动端适配：当前版本以 PC 端为主，手机端图表密度可进一步优化；', { size: 22 }),
  p('3. 空间维度：尚未加入 GIS 地图与热力图模块，下一版将引入；', { size: 22 }),
  p('4. 故事性：报告偏「数据呈现」，可进一步加入人物故事、保护区案例等叙事元素。', { size: 22 })
]);

const teamwork = sectionBlock('七、小组合作情况及任务分工', [
  p('本作品由 4 人小组合作完成，分工如下：', { size: 22, bold: true }),
  p('• 甘校松（P241012401）：数据检索与整理，原始表格清洗、湿地覆盖率公式计算与校对。', { size: 22 }),
  p('• 安欣（P241012428）：图表配色与交互设计，七大区域配色方案、动画细节、用户体验优化。', { size: 22 }),
  p('• 邓继诚（P241012460）：作品说明撰写、统筹，HTML/ECharts 编码、Word/PDF 报告生成、项目协调。', { size: 22 }),
  p('• 方子轩（P241012407）：数据校验与查漏补缺，跨源数据交叉验证，参考资料整理。', { size: 22 }),
  p('小组每周固定 2 次碰头会，使用腾讯会议 + 微信群协作；累计提交 git commit 50+ 次。', { size: 22 })
]);

const other = sectionBlock('八、其它需要说明的内容', [
  p('文案、数据来源及参考文献：', { bold: true, size: 24, color: DARK_GREEN }),
  p('[1] 国家林业和草原局. 第九次全国森林资源清查成果（2024）.', { size: 22 }),
  p('[2] 国家统计局. 中国第一次全国湿地资源调查（1995-2003）汇总数据.', { size: 22 }),
  p('[3] 自然资源部. 中国省域行政区划与面积基础地理数据.', { size: 22 }),
  p('[4] 联合国环境规划署（UNEP）.《全球湿地展望：湿地与水》（2021）. 内部视觉风格参考.', { size: 22 }),
  p('[5] Apache ECharts 5 可视化库官方文档：https://echarts.apache.org/zh/index.html', { size: 22 }),
  p('[6] 国务院新闻办公室.《中国的生物多样性保护》白皮书（2021）.', { size: 22 })
]);

const workURL = sectionBlock('九、作品访问地址及效果展示', [
  p('（一）作品访问地址：', { bold: true, size: 24, color: DARK_GREEN }),
  p(`网页版链接：${PROJECT_URL}`, { size: 22 }),
  p('本地路径：C:\\Users\\Night\\WorkBuddy\\2026-06-10-21-09-39\\china-forest-wetland.html', { size: 22 }),
  p('（二）作品效果展示（多角度、全方位展示作品）：', { bold: true, size: 24, color: DARK_GREEN }),
  p('① 顶部 Hero 区：渐变绿色 + 蓝色背景，左上角标题「中国森林、湿地覆盖率各省对比」，副标题「从东南绿岭到西北大漠——中国生态资源的空间分布格局」；', { size: 22 }),
  p('② Tab 1 总览：4 张数据卡片（全国森林覆盖率 25.00%、湿地覆盖率 4.01%、福建 66.80% 森林最高、上海 50.70% 湿地最高）+ 七大区域森林/湿地均值雷达图 + 七大区域森林/湿地覆盖率进度条；', { size: 22 }),
  p('③ Tab 2 排名对比：双轴柱状图（绿柱森林 + 蓝柱湿地），下方滑块可缩放；', { size: 22 }),
  p('④ Tab 3 省市检索：散点图（横轴森林覆盖率 × 纵轴湿地覆盖率，气泡大小=土地面积，颜色=七大分区），支持 brush 框选 + 搜索框 + 区域筛选 + 按字段排序的速查表；', { size: 22 }),
  p('⑤ Tab 4 双省对比：两个下拉框选择任意两省，雷达图 + 柱状图 + 自动生成对比小结文字。', { size: 22 }),
  p('⑥ 底部：数据来源、计算公式、更新时间三段署名信息。', { size: 22 })
]);

// Build document
const doc = new Document({
  creator: '2024级新闻学3班第9组',
  title: '中国森林、湿地覆盖率各省对比 - 作品说明书',
  description: '数据新闻理论与实践 期末作品',
  styles: {
    default: {
      document: { run: { font: { eastAsia: '宋体', ascii: '宋体', hAnsi: '宋体' } } }
    }
  },
  sections: [{
    properties: {
      page: { margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 } }
    },
    children: [
      // Top header
      new Paragraph({ spacing: { after: 200 }, children: [new TextRun('')] }),
      headerTable,
      new Paragraph({ spacing: { after: 100 }, children: [new TextRun('')] }),
      infoTable,
      new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: '《数据新闻理论与实践》作品创作说明', bold: true, size: 32, color: DARK_GREEN, font: { eastAsia: '黑体', ascii: '黑体', hAnsi: '黑体' } })], alignment: AlignmentType.CENTER }),
      membersTable,
      new Paragraph({ spacing: { after: 100 }, children: [new TextRun('')] }),
      ...workIntro,
      ...workLink,
      ...workNotes,
      ...toolDesc,
      ...materialDesc,
      ...highlights,
      ...teamwork,
      ...other,
      ...workURL,
      // Footer
      new Paragraph({ spacing: { before: 400 }, children: [new TextRun('')] }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
          bottom: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
          left: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
          right: { style: BorderStyle.SINGLE, size: 12, color: DARK_GREEN },
          insideHorizontal: { style: BorderStyle.NONE, size: 0 },
          insideVertical: { style: BorderStyle.NONE, size: 0 }
        },
        rows: [new TableRow({
          children: [new TableCell({
            width: { size: 100, type: WidthType.PERCENTAGE },
            children: [p('西北民族大学  新闻传播学院', { bold: true, size: 26, align: AlignmentType.CENTER, color: DARK_GREEN })]
          })]
        })]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  const outPath = process.argv[2] || 'C:\\Users\\Night\\WorkBuddy\\2026-06-10-21-09-39\\数据新闻作品说明书_中国森林湿地覆盖率各省对比.docx';
  fs.writeFileSync(outPath, buf);
  console.log('Saved:', outPath, 'Size:', buf.length, 'bytes');
});
