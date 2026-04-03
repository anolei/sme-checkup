#!/usr/bin/env python3
"""
申报书体检报告自我复核脚本（开源版）
严格按照 SKILL.md 阶段六的 14 项复核清单逐项检测。

使用方法：
    python3 self_review.py <体检报告.md路径>

返回：
    - 0: 全部通过
    - 1: 发现严重缺失 → agent 必须严格按照技能重新执行缺失步骤
"""

import sys
from pathlib import Path


def configure_output_streams():
    """避免 Windows 默认 GBK 控制台在输出 emoji 时触发 UnicodeEncodeError。"""
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(errors="replace")
            except ValueError:
                pass


configure_output_streams()


def check_content(report_path):
    """严格按照 SKILL.md 阶段六的 14 项清单逐项检测"""
    with open(report_path, 'r', encoding='utf-8') as f:
        content = f.read()

    issues = []    # 严重错误：必须重做
    warnings = []  # 警告：建议完善

    # ── 1. Y年是否正确识别 ──
    if 'Y=' not in content and '基准年' not in content:
        issues.append("[1] 未标注基准年（Y=XXXX年），无法确认年份推导是否正确")

    # ── 2. 红线指标合规底线是否全部输出并判定 ──
    indicator_keywords = ['营业收入', '主营业务收入', '资产负债率', '研发费用', '知识产权', '市场占有率']
    indicator_hits = sum(1 for kw in indicator_keywords if kw in content)
    if indicator_hits < 4:
        issues.append(f"[2] 红线指标覆盖不足（仅命中 {indicator_hits}/6 项关键词），请检查壹章表格是否完整")

    # ── 3. 全文逻辑一致性核查 ──
    if '逻辑' not in content and '一致性' not in content and '自洽' not in content:
        issues.append("[3] 缺少跨表单逻辑一致性核查（行业代码↔产品代码↔六基勾选）")

    # ── 4. 经济数据穿透式验算 ──
    if '穿透' not in content and '倒推' not in content and '复算' not in content:
        issues.append("[4] 缺少经济数据穿透式验算与填报一致性核查（贰章）")

    # ── 5. 知识产权深度核查 ──
    if '知识产权' not in content and '专利' not in content:
        issues.append("[5] 缺少知识产权深度核查")

    # ── 6. 双维度产业定位评判 ──
    if '维度一' not in content and '维度二' not in content and '双维度' not in content:
        issues.append("[6] 缺少双维度产业定位评判（维度一/维度二分析表格）")

    # ── 7. E-01 产业链配套找茬与改写骨架 ──
    if '产业链' not in content:
        issues.append("[7] 缺少产业链配套描述的诊断与改写建议")

    # ── 8. E-02 产学研缺失项提醒 ──
    if '产学研' not in content:
        issues.append("[8] 缺少产学研合作描述的诊断")

    # ── 9. E-03 企业简介四维度核查 ──
    dimension_keywords = ['深耕领域', '行业地位', '补短板', '产业链价值', '知识产权', '战略规划', '精细化管理']
    dimension_hits = sum(1 for kw in dimension_keywords if kw in content)
    if dimension_hits < 3:
        issues.append(f"[9] 企业简介四维度核查覆盖不足（命中 {dimension_hits}/7），A/B/C/D 维度可能遗漏")

    # ── 10. E-04 主导产品名称推导 + 级联一致性 ──
    if '主导产品' not in content and '产品名称' not in content:
        warnings.append("[10] 未发现主导产品名称的诊断与升维推导")

    # ── 11. 分级修改建议（🔴🟡🟢）──
    triage_keywords = ['优先修正', '审核警告', '优化建议', '🔴', '🟡', '🟢']
    triage_hits = sum(1 for kw in triage_keywords if kw in content)
    if triage_hits < 2:
        issues.append("[11] 缺少红黄绿三级分级修改建议（伍章）")

    # ── 12. 关键风险点 ──
    if '风险' not in content:
        issues.append("[12] 缺少关键风险点提示")

    # ── 13. 是否使用了报告模板格式 ──
    structure_keywords = ['壹', '贰', '叁', '肆', '伍', '陆']
    structure_hits = sum(1 for kw in structure_keywords if kw in content)
    if structure_hits < 5:
        warnings.append(f"[13] 报告结构可能未按模板输出（仅命中 {structure_hits}/6 个章节标记）")

    # ── 14. 免责声明 ──
    if '仅供' not in content and '风险提示' not in content and '不构成' not in content:
        issues.append("[14] 缺少 AI 免责声明")

    return issues, warnings


def print_report(issues, warnings):
    """打印复核报告"""
    print("=" * 60)
    print("申报书体检报告自我复核报告")
    print("=" * 60)

    if not issues and not warnings:
        print("\n✅ 全部 14 项检查通过！")
        print("体检报告完整，可以输出终稿。")
        return 0

    if issues:
        print(f"\n🔴 发现 {len(issues)} 个严重缺失（必须重做）：")
        for issue in issues:
            print(f"  ✗ {issue}")

    if warnings:
        print(f"\n🟡 发现 {len(warnings)} 个警告（建议完善）：")
        for warning in warnings:
            print(f"  △ {warning}")

    print("\n" + "=" * 60)

    if issues:
        print("❌ 复核未通过")
        print("→ agent 必须严格按照 SKILL.md 重新执行上述缺失步骤，")
        print("  补齐后重新生成报告并再次执行本脚本复核。")
        return 1
    else:
        print("⚠️ 复核通过（有警告）：建议完善警告项，不影响输出终稿。")
        return 0


def main():
    if len(sys.argv) < 2:
        print("用法：python3 self_review.py <体检报告.md路径>")
        sys.exit(1)

    report_path = sys.argv[1]

    if not Path(report_path).exists():
        print(f"错误：文件不存在 - {report_path}")
        sys.exit(1)

    issues, warnings = check_content(report_path)
    exit_code = print_report(issues, warnings)
    sys.exit(exit_code)


if __name__ == '__main__':
    main()
