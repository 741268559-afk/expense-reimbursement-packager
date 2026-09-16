# Expense Reimbursement Packager Skill

一个可供 Codex、WorkBuddy、Claude Code 及其他支持 Agent Skills 或终端脚本的智能体使用的财务报销整理工具包。它可以从支付截图、发票、行程单和 Excel 费用报销单模板中生成可复核、可打印、可归档的报销材料包。

项目默认适配中文报销流程和钉钉审批字段，但分类、发票要求、审批要求及费用报销单模板均可配置。

## 能力

- 从支付截图或文件名生成报销明细草稿，并分配稳定的 `FY001` 凭证编号。
- 按交通费、餐饮费、设备费、场地费、演员费和其他费用汇总，也支持自定义分类。
- 在提交审批前输出逐笔金额确认单、分类小计和总报销金额。
- 检查退款、拆分支付、个人费用、外币和多金额等异常。
- 检查审批字段是否完整，并要求审批金额与支付凭证总额精确一致。
- 分析原始发票和替代发票的覆盖金额、重复文件、号码、日期及购买方名称。
- 默认填充内置 Excel 费用报销单模板，也支持用户提供的其他模板。
- 在审批前、发票补齐后生成独立的钉钉提交包，包含提交版费用报销单、全部发票、行程单和发票打印 PDF。
- 生成明细表、三张一页截图、单张发票、一键打印 PDF、财务提交文件夹和审计摘要。
- 对金额、公式、发票覆盖、附件页数、连续页码和最终 ZIP 执行自动校验。

## Agent 与系统支持

- Agent：Codex、WorkBuddy 个人版/企业版、Claude Code，以及能够读取 `SKILL.md` 并执行 Python 的其他 Agent。
- 系统：Windows 10/11、macOS；核心 Python 流程也可在 Linux 使用。
- OCR：macOS Apple Vision，或 Windows/macOS Tesseract；没有 OCR 时仍可通过文件名和 Excel 审核流程使用。

发布页提供两个包：

- `universal`：采用保守的 Agent Skills frontmatter，适合 Codex、Claude Code 和其他兼容宿主。
- `workbuddy`：包含 WorkBuddy 要求的中英文字段、版本和作者信息，并包含企业版需要的 `manifest.yaml`。

## 安装

需要 Python 3.10 或更高版本。下载或克隆仓库后，使用同一个跨平台初始化命令：

```bash
git clone https://github.com/741268559-afk/expense-reimbursement-packager.git
cd expense-reimbursement-packager
python scripts/bootstrap.py
```

Windows PowerShell 可以把第一条命令写成 `py -3 scripts\bootstrap.py`。初始化结果会返回虚拟环境 Python 的完整路径。

### Codex 与其他 Agent Skills 宿主

把 universal ZIP 解压到宿主的 Skill 目录，或直接克隆到该目录。Codex 的示例位置是 `~/.codex/skills/expense-reimbursement-packager`，其他宿主使用各自的 Skill 目录。

### WorkBuddy

在 WorkBuddy 的“添加技能”中选择“上传技能”，导入 Release 中的 `expense-reimbursement-packager-workbuddy-*.zip`。企业版也可在 Skill 管理中上传同一个 ZIP。

WorkBuddy 官方导入规范见 [Skill 开发文档](https://open.workbuddy.cn/docs/skill)；平台操作见 [技能安装文档](https://www.workbuddy.cn/docs/workbuddy/From-Beginner-to-Expert-Guide/Function-Description/Skills-Market)。

## 快速验证

先运行合成数据自测。它不会读取真实报销文件：

```bash
python scripts/self_test.py --out expense-reimbursement-self-test
```

轻量环境检查：

```bash
python scripts/check_readiness.py --out reimbursement-readiness.json
```

仓库已内置经过校验的 `assets/费用报销单模板.xlsx` 和 `assets/form_cells.json`，正常安装时轻量检查应返回 `ready`。

## 基本用法

第一次运行先生成报销金额确认单、异常清单、发票缺口和审批字段缺口：

```bash
python3 scripts/package_from_folder.py \
  --input /path/to/expense-files \
  --project-name "项目名称" \
  --reimburser "报销人" \
  --out /path/to/output
```

如果发票不足，先补充替票并重新运行。生成提交版报销单前，还需要明确报销类型、领款人和开票对象；这些信息可通过部分 `--approval-metadata` 或复用配置提供。信息与发票都完整后，程序返回 `ready_for_dingtalk`，并生成：

```json
{
  "reimbursement_type": "项目报销",
  "payee": "领款人",
  "invoice_entity": "公司开票名称"
}
```

此时不填写 `dingding_number` 和 `approved_amount`，将该 JSON 通过 `--approval-metadata` 传入并重新运行即可生成提交包。

- `钉钉提交材料/`：提交版费用报销单、原始/替代发票、行程单和便捷打印 PDF
- `<项目>_<报销人>_钉钉提交材料.zip`：可用于钉钉附件上传的整包文件
- `dingtalk_submission_result.json`：上述文件的机器可读索引

此时费用报销单中的钉钉审批编号可以为空，因为编号尚未产生。

审批完成后，把截图中可见字段抄录到 JSON，再重新运行：

```json
{
  "dingding_number": "审批编号",
  "reimbursement_type": "项目报销",
  "reimburser": "报销人",
  "payee": "领款人",
  "invoice_entity": "公司开票名称",
  "approved_amount": 1234.56
}
```

```bash
python3 scripts/package_from_folder.py \
  --input /path/to/expense-files \
  --invoice-input /path/to/original-invoices \
  --replacement-invoice-input /path/to/approved-replacement-invoices \
  --project-name "项目名称" \
  --reimburser "报销人" \
  --approval-metadata /path/to/approval_metadata.json \
  --out /path/to/final-output
```

上述命令会自动使用内置费用报销单。只有需要替换为其他表样时，才传入 `--form-cells /path/to/form_cells.json` 或 `--form-template /path/to/template.xlsx`。

详细字段、模板接入和分阶段审核流程见 [SKILL.md](SKILL.md) 与 [references/workflow.md](references/workflow.md)。

Windows、macOS、OCR 和配置目录说明见 [references/platforms.md](references/platforms.md)。

## 构建发布包

```bash
python scripts/build_release_packages.py --out dist
```

该命令同时生成 universal ZIP、WorkBuddy ZIP 和 SHA-256 校验文件。

## 数据与合规边界

- 脚本默认在本地处理文件，不包含上传真实票据的代码。
- 不要把真实支付截图、身份证明、发票、审批截图、额外的公司模板或生成结果提交到公开仓库。`assets/费用报销单模板.xlsx` 是仓库所有者明确授权公开的内置表样。
- OCR 和文件解析结果必须人工复核，尤其是金额、日期、币种、退款状态和费用用途。
- 发票文件检查不等同于税务平台验真。真伪、可抵扣性和替代发票能否使用，应由财务或税务人员确认。
- 该项目提供材料整理与一致性校验，不构成财务、税务或法律意见。

## 仓库结构

```text
expense-reimbursement-packager/
|-- SKILL.md
|-- AGENTS.md
|-- CLAUDE.md
|-- manifest.yaml
|-- agents/openai.yaml
|-- adapters/workbuddy/
|-- scripts/
|-- references/
|-- assets/
|   |-- 费用报销单模板.xlsx
|   `-- form_cells.json
|-- requirements.txt
`-- LICENSE
```

## 许可证

[MIT](LICENSE)
