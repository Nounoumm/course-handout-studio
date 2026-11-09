// Course Handout Studio — student-facing handout generator.
//
// Tools:
//   handout_templates  list built-in handout templates
//   handout_scaffold   write a starter spec for a topic
//   handout_check      validate a spec and report structure/coverage issues
//   handout_build      render DOCX and convert it to PDF
//
// The heavy lifting (OOXML layout) lives in builder/handout_builder.py, run
// with the Python distribution shipped by the harness. PDF conversion uses the
// host's officeToPdf service, so no external LibreOffice lookup is needed.

import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { mkdir, readFile, stat, writeFile } from 'node:fs/promises'
import crypto from 'node:crypto'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const HERE = path.dirname(fileURLToPath(import.meta.url))
const BUILDER = path.join(HERE, 'builder', 'handout_builder.py')
const SKILL_FILE = path.join(HERE, 'skills', 'course-handout.md')

export const name = 'course-handout'

export const inject = ['tools', 'skills', 'officeToPdf', 'subprocess', 'workingDirectory']

// ---------------------------------------------------------------------------
// Minimal tool helper: the harness `defineTool` lives in a package that a
// linked workspace plugin cannot resolve, so convert the small parameter DSL
// used here into plain JSON Schema ourselves.
// ---------------------------------------------------------------------------

function fieldSchema(field) {
  const schema = {}
  if (field.type) schema.type = field.type
  if (field.description) schema.description = field.description
  if (field.enum) schema.enum = field.enum
  if (field.type === 'object' && field.properties) {
    schema.properties = {}
    const required = []
    for (const [key, value] of Object.entries(field.properties)) {
      schema.properties[key] = fieldSchema(value)
      if (value.required) required.push(key)
    }
    if (required.length) schema.required = required
    schema.additionalProperties = false
  }
  if (field.items) schema.items = fieldSchema(field.items)
  return schema
}

function defineTool(definition) {
  const { parameters, ...rest } = definition
  if (!parameters) return { ...rest, parameters: { type: 'object', properties: {} } }
  const schema = fieldSchema({ type: 'object', properties: parameters })
  return { ...rest, parameters: schema }
}

function statsSchema() {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['modules', 'chapters', 'sections', 'blocks', 'questions', 'chars', 'code_lines', 'est_pages'],
    properties: {
      modules: { type: 'integer' },
      chapters: { type: 'integer' },
      sections: { type: 'integer' },
      blocks: { type: 'integer' },
      questions: { type: 'integer' },
      chars: { type: 'integer' },
      code_lines: { type: 'integer' },
      est_pages: { type: 'integer' },
    },
  }
}


// ---------------------------------------------------------------------------
// Templates
// ---------------------------------------------------------------------------

const TEMPLATES = {
  standard: {
    name: '标准教材型',
    description: '模块—章—节三级结构；每章含学习目标、先修知识、概念与例题、常见错误、小结和分层习题。',
    bestFor: '系统学习一门课，或从零补基础',
    chapterShape: [
      ['导论：这门课学什么', ['问题引入', '知识地图', '学习方法']],
      ['核心概念与表示', ['概念定义', '表示方法', '基本性质']],
      ['基本方法与推导', ['方法一', '方法二', '对比与选择']],
      ['综合应用与建模', ['典型应用', '综合例题', '常见错误']],
    ],
  },
  review: {
    name: '考前复习型',
    description: '以知识地图、重点速览、题型归纳和模拟自测为主，压缩推导，突出考点与易错点。',
    bestFor: '期末复习、考研复盘、面试前快速回顾',
    chapterShape: [
      ['知识地图与考点分布', ['全章框架', '考点权重', '复习顺序']],
      ['重点概念速览', ['必背定义', '关键公式', '适用条件']],
      ['题型归纳与模板', ['题型一', '题型二', '解题模板']],
      ['易错点与模拟自测', ['高频错误', '限时练习', '自查清单']],
    ],
  },
  project: {
    name: '项目驱动型',
    description: '从一个真实任务出发，按“情境—拆解—知识补给—实现—复盘”组织，知识在解决问题中出现。',
    bestFor: '项目式学习、课程设计、上手实践',
    chapterShape: [
      ['项目情境与目标', ['任务描述', '验收标准', '环境准备']],
      ['任务拆解与技术路线', ['子任务', '依赖关系', '里程碑']],
      ['知识补给站', ['够用的概念', '关键 API', '最小示例']],
      ['实现、调试与复盘', ['实现步骤', '调试方法', '复盘与扩展']],
    ],
  },
  bilingual: {
    name: '双语对照型',
    description: '中文讲解为主，关键术语给出英文对照，并用外文教材常见的图示、直觉解释和“为什么”组织内容。',
    bestFor: '阅读英文教材、准备双语课程或出国考试',
    chapterShape: [
      ['概念与术语对照', ['中文定义', 'English term', '常见译法辨析']],
      ['原理图解', ['直观图示', '逐步推导', '几何/物理意义']],
      ['例题精讲', ['基础例题', '综合例题', '英文题干示例']],
      ['术语表与练习', ['核心术语', '中英互译练习', '拓展阅读']],
    ],
  },
}

function templateInfo(id) {
  const tpl = TEMPLATES[id] || TEMPLATES.standard
  return {
    id: TEMPLATES[id] ? id : 'standard',
    name: tpl.name,
    description: tpl.description,
    bestFor: tpl.bestFor,
    chapters: tpl.chapterShape.length,
  }
}

function scaffoldSpec({ topic, audience, template, chapters, exercises, language }) {
  const tplId = TEMPLATES[template] ? template : 'standard'
  const tpl = TEMPLATES[tplId]
  const count = Math.max(3, Math.min(12, Number(chapters) || 6))
  const shape = []
  for (let i = 0; i < count; i++) {
    const seed = tpl.chapterShape[i % tpl.chapterShape.length]
    const round = Math.floor(i / tpl.chapterShape.length)
    const suffix = round > 0 ? `（${round + 1}）` : ''
    shape.push([seed[0] + suffix, seed[1]])
  }

  const modules = []
  const perModule = count <= 4 ? count : Math.ceil(count / Math.min(3, Math.ceil(count / 2)))
  for (let start = 0; start < shape.length; start += perModule) {
    const slice = shape.slice(start, start + perModule)
    const mIndex = modules.length + 1
    modules.push({
      title: mIndex === 1 ? '第一部分　基础与直觉' : `第 ${mIndex} 部分　方法与进阶`,
      description: '用一两句话说明本模块要解决的问题，以及它与前后模块的关系。',
      chapters: slice.map(([title, sections], ci) => ({
        title: `第 ${start + ci + 1} 章　${title}`,
        objectives: [
          `理解${topic}中“${title}”的核心问题`,
          '能够独立完成本章的基础例题',
          '能够指出本章最常见的两个错误',
        ],
        prerequisites: start + ci === 0 ? ['无（零基础起点）'] : ['上一章的核心结论'],
        sections: sections.map((s, si) => ({
          title: `${start + ci + 1}.${si + 1}　${s}`,
          body:
            '（待补充）在这里写本节正文。\n\n' +
            '建议顺序：先用一个具体问题引出，再给定义与记号，然后推导或解释“为什么”，' +
            '最后给 1–2 道例题。\n\n' +
            '> 提示：支持 Markdown 段落、**粗体**、`行内代码`、列表、表格、代码块和引用块。',
        })),
        summary: ['本章最重要的结论', '一个容易混淆的点', '与下一章的衔接'],
        exercises: exercises === 'none' ? {} : {
          basic: ['基础题：直接套用定义或公式'],
          intermediate: ['提高题：需要两步以上的推理'],
          ...(exercises === 'full' ? { challenge: ['挑战题：综合应用或开放问题'] } : {}),
        },
        answers: exercises === 'none'
          ? []
          : ['参考答案（基础题）', '参考答案（提高题）', ...(exercises === 'full' ? ['参考答案（挑战题）'] : [])],
      })),
    })
  }

  return {
    schema: 'course-handout/1',
    meta: {
      title: topic,
      subtitle: '面向学生的模块化讲义',
      series: '课程讲义',
      author: '课程讲师',
      date: new Date().toISOString().slice(0, 7).replace('-', ' 年 ') + ' 月',
      audience: audience || '本课程学习者',
      language: language || 'zh-CN',
      template: tplId,
      include_toc: true,
      include_answers: exercises !== 'none',
    },
    preface: '',
    modules,
    appendix: {
      glossary: [['术语', '用一句话解释这个术语']],
      references: [
        {
          title: '配套视频或公开课（请替换为实际链接）',
          url: 'https://example.com',
          note: '说明它对应哪一章、为什么推荐',
        },
      ],
    },
  }
}

// ---------------------------------------------------------------------------
// Python discovery and execution
// ---------------------------------------------------------------------------

function pythonCandidates(config) {
  const exe = process.platform === 'win32' ? 'python.exe' : 'python3'
  const candidates = []
  if (config && config.pythonPath) candidates.push(config.pythonPath)
  if (process.env.DSH_PYTHON) candidates.push(process.env.DSH_PYTHON)

  const home = process.env.DSH_HOME || path.join(os.homedir(), '.dsh')
  const runtimes = path.join(home, 'dsh-runtimes')
  try {
    for (const entry of readdirSync(runtimes)) {
      candidates.push(path.join(runtimes, entry, 'dependencies', 'python', exe))
    }
  } catch { /* no profile runtimes */ }

  const resources = path.resolve(path.dirname(process.execPath), 'resources', 'runtime')
  try {
    for (const entry of readdirSync(resources)) {
      candidates.push(path.join(resources, entry, 'dependencies', 'python', exe))
    }
  } catch { /* not an Electron layout */ }

  candidates.push(exe, 'python3', 'python')
  return candidates
}

function makePythonResolver(ctx, config) {
  let pending
  return () => {
    if (!pending) {
      pending = (async () => {
        for (const candidate of pythonCandidates(config)) {
          try {
            return await ctx.subprocess.resolveExecutable(candidate)
          } catch { /* try next */ }
        }
        throw new Error(
          '找不到可用的 Python。请确认 harness 运行时完整，或在插件配置里设置 pythonPath。',
        )
      })()
    }
    return pending
  }
}

async function runBuilder(ctx, python, args, cwd, signal) {
  const handle = ctx.subprocess.spawn({
    argv: [python, BUILDER, ...args],
    cwd,
    stdio: {
      stdin: 'ignore',
      stdout: { maxBytes: 4_000_000 },
      stderr: { maxBytes: 400_000 },
    },
    graceMs: 5000,
    signal,
  })
  const outcome = await handle.done
  const stdout = handle.collected.stdout?.readFrom(0).text ?? ''
  const stderr = handle.collected.stderr?.readFrom(0).text ?? ''
  if (outcome.exitCode !== 0) {
    let detail = stdout.trim()
    try {
      const parsed = JSON.parse(detail.split('\n').filter(Boolean).at(-1))
      if (parsed?.error) detail = parsed.error
    } catch { /* keep raw */ }
    const reason = detail || stderr.trim().split('\n').slice(-3).join('\n') || `exit ${outcome.exitCode}`
    throw new Error(reason)
  }
  const last = stdout.split('\n').filter((line) => line.trim().startsWith('{')).at(-1)
  if (!last) throw new Error(`讲义生成器没有返回结果：${stderr.trim().slice(-500)}`)
  return JSON.parse(last)
}

// ---------------------------------------------------------------------------
// File helpers
// ---------------------------------------------------------------------------

function resolveFrom(cwd, p) {
  return path.isAbsolute(p) ? p : path.resolve(cwd, p)
}

function displayPath(cwd, abs) {
  const rel = path.relative(cwd, abs)
  return rel && !rel.startsWith('..') ? rel : abs
}

function slugify(title) {
  const cleaned = String(title || '')
    .replace(/[\\/:*?"<>|\r\n\t]/g, ' ')
    .replace(/\s+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 64)
  return cleaned || `handout-${Date.now()}`
}

async function ensureDir(dir) {
  await mkdir(dir, { recursive: true })
}

function uniquePath(base) {
  if (!existsSync(base)) return base
  const ext = path.extname(base)
  const stem = base.slice(0, -ext.length)
  for (let i = 2; i < 1000; i++) {
    const candidate = `${stem}-${i}${ext}`
    if (!existsSync(candidate)) return candidate
  }
  return `${stem}-${Date.now()}${ext}`
}

// ---------------------------------------------------------------------------
// Plugin
// ---------------------------------------------------------------------------

export function apply(ctx, config = {}) {
  const resolvePython = makePythonResolver(ctx, config)
  const outputSubdir = config.outputSubdir || 'handouts'

  const currentDir = async (exec) => {
    if (exec.agent) return ctx.workingDirectory.ensure(exec.agent, exec.signal)
    return process.cwd()
  }

  // -- templates -----------------------------------------------------
  ctx.effect(() => ctx.tools.register(defineTool({
    name: 'handout_templates',
    description: '列出可用的讲义模板（标准教材型、考前复习型、项目驱动型、双语对照型），'
      + '包含每个模板的适用场景与章节骨架。学生没有指定风格时，用它给出选项。',
    parameters: {},
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        required: ['templates'],
        properties: {
          templates: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              required: ['id', 'name', 'description', 'bestFor', 'chapters'],
              properties: {
                id: { type: 'string' },
                name: { type: 'string' },
                description: { type: 'string' },
                bestFor: { type: 'string' },
                chapters: { type: 'integer' },
              },
            },
          },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: value.templates
          .map((t) => `- ${t.name}（${t.id}）：${t.description} 适合：${t.bestFor}；骨架 ${t.chapters} 章。`)
          .join('\n'),
      }],
    },
    execute() {
      return { templates: Object.keys(TEMPLATES).map((id) => templateInfo(id)) }
    },
  })))

  // -- scaffold ------------------------------------------------------
  ctx.effect(() => ctx.tools.register(defineTool({
    name: 'handout_scaffold',
    description: '为一个主题写出讲义规格骨架（JSON 文件），包含模块、章节、学习目标、小节占位和习题框架。'
      + '随后用编辑工具补全内容，再调用 handout_build。骨架里的“（待补充）”必须替换为真实内容。',
    parameters: {
      topic: {
        type: 'string',
        required: true,
        description: '讲义主题，例如“高中物理：牛顿运动定律”或“Python 列表与字典”。',
      },
      audience: {
        type: 'string',
        description: '适用对象，例如“高一学生”“零基础转行者”“期末复习的本科生”。',
      },
      template: {
        type: 'string',
        enum: Object.keys(TEMPLATES),
        description: '模板 id；默认 standard。先调用 handout_templates 了解差异。',
      },
      chapters: {
        type: 'integer',
        description: '章节数量，3–12，默认 6。按知识点自然分组，不要为凑数拆章。',
      },
      exercises: {
        type: 'string',
        enum: ['none', 'basic', 'full'],
        description: '习题规模：none 不带题，basic 基础+提高，full 再加挑战题。默认 basic。',
      },
      output_path: {
        type: 'string',
        description: '规格文件路径，默认 handouts/<主题>.json。',
      },
      overwrite: {
        type: 'boolean',
        description: '目标文件已存在时是否覆盖，默认 false。',
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        required: ['spec_path', 'template', 'chapters', 'next'],
        properties: {
          spec_path: { type: 'string' },
          template: { type: 'string' },
          chapters: { type: 'integer' },
          next: { type: 'string' },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `已生成规格骨架：${value.spec_path}（模板 ${value.template}，${value.chapters} 章）。\n${value.next}`,
      }],
    },
    async execute(args, exec) {
      const cwd = await currentDir(exec)
      const topic = String(args.topic || '').trim()
      if (!topic) throw new Error('topic 不能为空')
      const target = resolveFrom(
        cwd,
        args.output_path || path.join(outputSubdir, `${slugify(topic)}.json`),
      )
      if (existsSync(target) && !args.overwrite) {
        throw new Error(`规格文件已存在：${displayPath(cwd, target)}。换一个路径，或传 overwrite=true。`)
      }
      const spec = scaffoldSpec({
        topic,
        audience: args.audience,
        template: args.template,
        chapters: args.chapters,
        exercises: args.exercises || 'basic',
        language: 'zh-CN',
      })
      await ensureDir(path.dirname(target))
      await writeFile(target, JSON.stringify(spec, null, 2), 'utf8')
      return {
        spec_path: displayPath(cwd, target),
        template: spec.meta.template,
        chapters: Number(args.chapters) || 6,
        next: '用编辑工具补全每个小节的 body、学习目标、小结和习题答案，然后运行 handout_check，再运行 handout_build。',
      }
    },
  })))

  // -- check ---------------------------------------------------------
  ctx.effect(() => ctx.tools.register(defineTool({
    name: 'handout_check',
    description: '检查讲义规格：章节结构、空小节、学习目标与小结、习题与答案数量、篇幅估计。'
      + '返回 error/warning 列表；有 error 时 handout_build 会拒绝生成（除非 force）。',
    parameters: {
      spec_path: {
        type: 'string',
        required: true,
        description: 'handout_scaffold 生成或手写的规格 JSON 路径。',
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        required: ['ok', 'stats', 'issues'],
        properties: {
          ok: { type: 'boolean' },
          stats: statsSchema(),
          issues: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              required: ['level', 'path', 'message'],
              properties: {
                level: { type: 'string' },
                path: { type: 'string' },
                message: { type: 'string' },
              },
            },
          },
        },
      },
      render: (_args, value) => {
        const lines = value.issues.map((i) => `- [${i.level}] ${i.path}: ${i.message}`)
        const head = value.ok ? '检查通过' : '检查未通过'
        return [{
          type: 'text',
          text: `${head}：${value.stats.chapters} 章 / ${value.stats.sections} 节 / `
            + `${value.stats.questions} 题 / 约 ${value.stats.est_pages} 页。\n`
            + (lines.length ? lines.join('\n') : '没有发现问题。'),
        }]
      },
    },
    async execute(args, exec) {
      const cwd = await currentDir(exec)
      const specPath = resolveFrom(cwd, args.spec_path)
      await stat(specPath)
      const python = await resolvePython()
      const result = await runBuilder(ctx, python, ['check', '--spec', specPath], cwd, exec.signal)
      return {
        ok: Boolean(result.ok),
        stats: result.stats,
        issues: result.issues || [],
      }
    },
  })))

  // -- build ---------------------------------------------------------
  ctx.effect(() => ctx.tools.register(defineTool({
    name: 'handout_build',
    description: '把讲义规格渲染成 DOCX，并转换成 PDF（默认同时产出）。返回两个文件路径、'
      + '结构统计和警告；生成后用 present 把文件交给学生。',
    parameters: {
      spec_path: {
        type: 'string',
        required: true,
        description: '讲义规格 JSON 路径。',
      },
      output_dir: {
        type: 'string',
        description: '输出目录，默认 handouts/。',
      },
      file_stem: {
        type: 'string',
        description: '文件名（不含扩展名），默认取讲义标题。',
      },
      make_pdf: {
        type: 'boolean',
        description: '是否转换 PDF，默认 true。',
      },
      force: {
        type: 'boolean',
        description: '即使检查有 error 也继续生成，默认 false。',
      },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        required: ['docx_path', 'stats', 'warnings'],
        properties: {
          docx_path: { type: 'string' },
          pdf_path: { type: 'string' },
          stats: statsSchema(),
          warnings: { type: 'array', items: { type: 'string' } },
        },
      },
      render: (_args, value) => [{
        type: 'text',
        text: `讲义已生成：\n- DOCX：${value.docx_path}\n`
          + (value.pdf_path ? `- PDF：${value.pdf_path}\n` : '- PDF：未生成\n')
          + `结构：${value.stats.chapters} 章 / ${value.stats.sections} 节 / ${value.stats.questions} 题。`
          + (value.warnings.length ? `\n警告：\n${value.warnings.map((w) => `- ${w}`).join('\n')}` : ''),
      }],
    },
    async execute(args, exec) {
      const cwd = await currentDir(exec)
      const specPath = resolveFrom(cwd, args.spec_path)
      const specText = await readFile(specPath, 'utf8')
      if (specText.length > 4_000_000) throw new Error('规格文件过大（超过 4 MB）')
      let spec
      try {
        spec = JSON.parse(specText)
      } catch (error) {
        throw new Error(`规格不是合法 JSON：${error.message}`)
      }
      const title = spec?.meta?.title || '讲义'
      const stem = slugify(args.file_stem || title)
      const outDir = resolveFrom(cwd, args.output_dir || outputSubdir)
      await ensureDir(outDir)
      let docxPath = path.join(outDir, `${stem}.docx`)
      if (!args.force) docxPath = uniquePath(docxPath)

      const python = await resolvePython()
      const buildArgs = ['build', '--spec', specPath, '--out', docxPath]
      if (args.force) buildArgs.push('--force')
      const result = await runBuilder(ctx, python, buildArgs, cwd, exec.signal)

      const warnings = (result.issues || [])
        .filter((issue) => issue.level !== 'error')
        .map((issue) => `${issue.path}: ${issue.message}`)

      let pdfPath
      if (args.make_pdf !== false) {
        try {
          const bytes = await readFile(docxPath)
          const hash = crypto.createHash('sha256').update(bytes).digest('hex')
          const converted = await ctx.officeToPdf.convert({
            extension: 'docx',
            priority: 'foreground',
            source: {
              key: `course-handout-${hash.slice(0, 32)}`,
              version: hash,
              bytes: bytes.byteLength,
              read: async () => ({ bytes, version: hash }),
            },
          }, exec.signal)
          pdfPath = docxPath.replace(/\.docx$/i, '.pdf')
          await writeFile(pdfPath, converted.pdf)
        } catch (error) {
          warnings.push(`PDF 转换失败：${error.message}`)
          pdfPath = undefined
        }
      }

      const stats = {
        modules: result.stats.modules,
        chapters: result.stats.chapters,
        sections: result.stats.sections,
        blocks: result.stats.blocks,
        questions: result.stats.questions,
        chars: result.stats.chars,
        code_lines: result.stats.code_lines,
        est_pages: result.stats.est_pages,
      }
      return {
        docx_path: displayPath(cwd, docxPath),
        ...(pdfPath ? { pdf_path: displayPath(cwd, pdfPath) } : {}),
        stats,
        warnings,
      }
    },
  })))

  // -- skill ---------------------------------------------------------
  if (existsSync(SKILL_FILE)) {
    const content = readFileSync(SKILL_FILE, 'utf8')
    ctx.effect(() => ctx.skills.register({
      name: 'course-handout',
      description: '面向学生的课程讲义生成：把一次提问扩展成模块化、知识点齐全的讲义，按需配习题与答案，'
        + '输出 DOCX 与 PDF。',
      whenToUse: '学生要求“生成/整理一份讲义、教程、复习资料、学习笔记”，或想把某个主题系统讲清楚时。',
      content,
      source: 'custom',
      invocation: { modelInvocable: true, userInvocable: true },
    }))
  }
}
