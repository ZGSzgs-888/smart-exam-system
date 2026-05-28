# 智能题库与自适应组卷系统 — 设计规格

## 1. 项目概述

### 1.1 项目定位
基于前后端分离架构的智能题库与自适应组卷系统，面向本科毕设与研究生复试项目展示。核心价值在于从"经验驱动"到"数据驱动"的转变，实现因材施教。

### 1.2 核心理念
- **智能题库：** 不只是存题，自动累积正确率、区分度等统计数据，越用越"聪明"
- **自适应组卷：** 基于学生能力模型，为每个学生生成最合适的试卷
- **数据闭环：** 考试数据反馈优化题库标签和学生画像，形成正循环

## 2. 技术栈

| 层级 | 技术选型 |
|------|---------|
| 前端 | Vue 3 + Element Plus + Vite |
| 后端 | Spring Boot + MyBatis-Plus |
| 数据库 | MySQL |
| 架构 | 前后端分离，RESTful API |
| 认证 | JWT |

## 3. 系统架构

```
前端 (Vue 3 + Element Plus)
  ├── 管理员端 ─── 全局管理视图
  ├── 教师端   ─── 题库管理、组卷、考试、学情
  └── 学生端   ─── 考试、成绩、错题、分析
        │
        │ REST API (JSON + JWT)
        ▼
后端 (Spring Boot)
  ├── 用户模块
  ├── 题库模块
  ├── 知识点模块
  ├── 组卷引擎（核心算法）
  ├── 考试模块
  ├── 成绩模块
  └── 统计模块
        │
        ▼
MySQL 数据库 (10 张核心表 + 扩展表)
```

## 4. 功能模块

### 4.1 账户管理
- 用户登录/登出（JWT）
- 三种角色：管理员 / 教师 / 学生
- 基于角色的菜单与权限控制
- 用户管理（管理员可增删改查用户）

### 4.2 科目管理
- 科目列表
- 科目增删改（管理员维护）

### 4.3 题库管理 ★
- 题目列表（分页、按科目/知识点/难度/题型筛选）
- 题目创建（支持单选/多选/填空/解答四种题型）
- 题目编辑/删除
- 富文本编辑（解答题图文混排）
- **Excel 批量导入**
- **自动统计：** 每道题自动积累正确率、区分度
- **智能辅助：** 推荐知识点、自动估算难度

### 4.4 知识点管理
- 树形结构展示（Element Plus Tree）
- 知识点增删改
- 拖拽调整层级
- 题目-知识点关联

### 4.5 智能组卷 ★★
- 组卷策略配置：科目、知识点范围、难度分布、题型与数量
- **三层组卷算法：** 约束过滤 → 均衡分配 → 能力适配
- 试卷预览
- 手动微调（替换/删除/添加题目）
- 保存试卷

### 4.6 在线考试
- 考试列表（待考/已考）
- 在线答题：计时、答题导航、题目切换
- 客观题自动判分
- 主观题教师批改
- 提交试卷

### 4.7 成绩管理
- 成绩列表
- 成绩详情（每题得分）
- 成绩导出 Excel
- 主观题批改

### 4.8 统计分析 ★
- 成绩分布柱状图
- 知识点掌握雷达图
- 个人历次成绩折线图
- 薄弱知识点排名
- 班级整体学情概览（平均分、及格率、优秀率）

### 4.9 自适应测评 ★★
- **学生能力模型：** 递推更新知识点掌握度
- 个性化错题本
- 薄弱知识点针对性推荐练习

## 5. 核心算法设计

### 5.1 智能组卷算法（三层架构）

**第一层：约束过滤**
```
输入：科目、题型数组、知识点范围
过程：从题库过滤符合条件的候选题目池
```

**第二层：均衡分配**
```
按"知识点"分组 → 每组按"难度"分层
→ 从各组按比例抽取，确保知识点覆盖和难度分布
```

**第三层：能力适配（自适应）**
```
查询学生知识点掌握度(t_mastery)
掌握度低的 → 对应知识点出简单/中等题
掌握度高的 → 对应知识点出中等/困难题
同一份试卷不同学生题目重合率 ≥ 60%
```

### 5.2 学生能力模型

```
掌握度更新公式：
  new_score = old_score × 0.7 + current_accuracy × 0.3

数据表：t_mastery (student_id, knowledge_id, mastery_score, ...)
```

### 5.3 统计分析

- 错题率聚合 → 定位薄弱知识点
- 成绩分布计算
- 趋势分析（移动平均）

## 6. 数据库设计

### 6.1 核心表（已有）

| 表名 | 说明 |
|------|------|
| t_user | 用户表（角色：管理员/教师/学生） |
| t_subject | 科目表 |
| t_knowledge | 知识点表（树形结构，parent_id） |
| t_question | 题目表 |
| t_question_option | 选项表（选择题专用） |
| t_paper | 试卷表（含组卷策略JSON） |
| t_paper_question | 试卷-题目关联表 |
| t_exam | 考试安排表 |
| t_answer_record | 答卷记录表 |
| t_answer_detail | 答题详情表 |

### 6.2 扩展表（新增）

| 表名 | 说明 |
|------|------|
| t_mastery | 学生知识点掌握度 |
| t_question_knowledge | 题目-知识点关联（多对多） |

### 6.3 字段优化

- t_question: 新增 `discrimination`（区分度）、`correct_rate`（正确率）
- t_exam: 新增 `strategy_json`（组卷策略快照）

## 7. 前端路由与页面

### 路由结构

```
/login                    → 登录
/admin/users              → 用户管理
/admin/subjects           → 科目管理
/admin/questions          → 所有题目（查看/删除）
/teacher/questions        → 题库管理（完整CRUD）
/teacher/knowledge        → 知识点管理
/teacher/paper            → 智能组卷
/teacher/paper/create     → 组卷配置
/teacher/paper/:id/preview → 试卷预览
/teacher/exams            → 考试管理
/teacher/scores           → 成绩管理
/teacher/statistics       → 学情分析
/student/exams            → 考试列表
/student/exams/:id        → 在线答题
/student/scores           → 成绩查看
/student/wrong-questions  → 错题本
/student/analysis         → 个人薄弱分析
```

### 核心页面

**在线答题页：** 左侧答题导航 + 右侧题目区 + 顶部计时器 + 底部进度条
**学情分析页：** 成绩分布柱状图 + 知识点雷达图 + 薄弱知识点排名
**组卷配置页：** 约束条件表单 + 预览区 + 手动微调

## 8. 后端模块结构

```
exam-system-backend/
├── common/         config, constant, exception, result
├── module/
│   ├── user/
│   ├── subject/
│   ├── question/
│   ├── knowledge/
│   ├── paper/
│   │   └── engine/        PaperGenerator, DifficultyCalculator, KnowledgeSelector
│   ├── exam/
│   ├── score/
│   └── statistics/        Charts data assembly
├── interceptor/   JWT + 角色权限
└── config/        跨域, Swagger
```

## 9. IRT 升级预留

- t_question 预留 difficulty_param, discrimination_param, guess_param 字段
- paper/engine/ 目录结构支持替换实现（PaperGenerator 接口化）
- 新增 IRTEngine.java 不影响上层

## 10. 数据闭环

```
学生能力 → 自适应组卷 → 考试 → 成绩分析
    ↑                            │
    └──────── 更新模型 ←─────────┘
```
