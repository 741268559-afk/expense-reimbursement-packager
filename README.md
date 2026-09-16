# Expense Reimbursement Packager Skill

一个面向 Codex 的财务报销整理 Skill。它可以从支付截图、发票、行程单和 Excel 费用报销单模板中生成可复核、可打印、可归档的报销材料包。

项目默认适配中文报销流程和钉钉审批字段，但分类、发票要求、审批要求及费用报销单模板均可配置。

## 能力

- 从支付截图或文件名生成报销明细草稿，并分配稳定的 `FY001` 凭证编号。
- 按交通费、餐饮费、设备费、场地费、演员费和其他费用汇总，也支持自定义分类。
- 在提交审批前输出逐笔金额确认单、分类小计和总报销金额。
- 检查退款、拆分支付、个人费用、外币和多金额等异常。
- 检查审批字段是否完整，并要求审批金额与支付凭证总额精确一致。
- 分析原始发票和替代发票的覆盖金额、重复文件、号码、日期及购买方名称。
- 填充用户提供的 Excel 费用报销单模板。
- 生成明细表、三张一页截图、单张发票、一键打印 PDF、财务提交文件夹和审计摘要。
- 对金额、公式、发票覆盖、附件页数、连续页码和最终 ZIP 执行自动校验。

## 安装

需要 Python 3.10 或更高版本。

```bash
git clone <repository-url> ~/.codex/skills/expense-reimbursement-packager
python3 -m pip install -r ~/.codex/skills/expense-reimbursement-packager/requirements.txt
```

重新打开 Codex 后，可通过 `$expense-reimbursement-packager` 显式调用；符合描述的报销任务也可以自动触发该 Skill。

macOS 上如已安装 Swift，Skill 可以使用 Apple Vision 做本地 OCR。其他平台仍可使用文件名解析、手工 Manifest 或 Excel 审核流程。

## 快速验证

先运行合成数据自测。它不会读取真实报销文件：

```bash
cd ~/.codex/skills/expense-reimbursement-packager
python3 scripts/self_test.py --out /tmp/expense-reimbursement-self-test
```

轻量环境检查：

```bash
python3 scripts/check_readiness.py --out /tmp/reimbursement-readiness.json
```

未配置费用报销单模板时，轻量检查返回 `needs_form_template` 是预期状态；明细表和打印材料仍可生成。

## 基本用法

第一次运行先生成报销金额确认单、异常清单、发票缺口和审批字段缺口：

```bash
python3 scripts/package_from_folder.py \
  --input /path/to/expense-files \
  --project-name "项目名称" \
  --reimburser "报销人" \
  --out /path/to/output
```

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
  --form-cells /path/to/form_cells.json \
  --out /path/to/final-output
```

详细字段、模板接入和两阶段审核流程见 [SKILL.md](SKILL.md) 与 [references/workflow.md](references/workflow.md)。

## 数据与合规边界

- 脚本默认在本地处理文件，不包含上传真实票据的代码。
- 不要把真实支付截图、身份证明、发票、审批截图、公司模板或生成结果提交到公开仓库。
- OCR 和文件解析结果必须人工复核，尤其是金额、日期、币种、退款状态和费用用途。
- 发票文件检查不等同于税务平台验真。真伪、可抵扣性和替代发票能否使用，应由财务或税务人员确认。
- 该项目提供材料整理与一致性校验，不构成财务、税务或法律意见。

## 仓库结构

```text
expense-reimbursement-packager/
|-- SKILL.md
|-- agents/openai.yaml
|-- scripts/
|-- references/workflow.md
|-- assets/
|-- requirements.txt
`-- LICENSE
```

## 许可证

[MIT](LICENSE)
