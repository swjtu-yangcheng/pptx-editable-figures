# -*- coding: utf-8 -*-
"""
一键示例：生成 3 页可编辑插图 → 几何自检 → 字符画预览。

用法：
    python examples/run_demo.py

产出：
    examples/demo.pptx     3 页（流程图 / 体系图 / 甘特图），原生形状、可编辑

非 Windows 或未安装 PowerPoint 时，字符画预览会自动跳过，其余流程正常。
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "scripts"))
sys.path.insert(0, os.path.join(ROOT, "templates"))

from figures import make_flowchart, make_system_diagram, make_gantt, merge_figures  # noqa: E402
from pptxfig import verify  # noqa: E402


def build(out_path):
    # ---------- 1. 流程图（含闭环回流）----------
    f1 = make_flowchart(
        [("信息采集", "学业数据、行为数据\n建立学生学业档案"),
         ("分级预警", "黄/橙/红三级\n教务初筛、辅导员复核"),
         ("预警谈话", "48 小时内约谈\n共同制定改进计划"),
         ("多方联动", "家长告知、教师联动\n组建帮扶小组"),
         ("跟踪销号", "月度跟进、定期会商\n达标后解除预警")],
        title="图2-1  学业预警与帮扶工作流程图",
        loop_from=4, loop_to=1,
        loop_label="未改善：升级干预措施，必要时转介专业力量",
        note="闭环管理：预警不解除不销号，销号后继续跟踪一个学期")

    # ---------- 2. 体系图（三层汇聚 + 回馈闭环）----------
    f2 = make_system_diagram(
        [("物质帮助", "奖助勤贷免补\n保障基本就学", "#D6E6F2"),
         ("能力拓展", "勤工助学、技能培训\n访学交流、就业支持", "#BBD6EA"),
         ("精神激励", "励志教育\n感恩诚信教育", "#D6E6F2")],
        "解困—育人—成才—回馈",
        [("价值引领", "志愿服务\n资助宣传大使", "#BBD6EA"),
         ("精准认定", "定量 + 定性\n动态调整", "#D6E6F2"),
         ("隐私保护", "去标识化公示\n限定数据用途", "#BBD6EA")],
        title="图8-1  发展型资助育人体系构成",
        loop_pair=(0, 0), loop_label="回馈反哺",
        note="说明：上层为资助途径，中层为育人目标链条，下层为支撑保障机制。")

    # ---------- 3. 甘特图（自动时间刻度 + 里程碑）----------
    # 注意：每项第 3 个数是【持续月数】，不是结束月序
    f3 = make_gantt(
        [("一  大纲论证与体例统一", 0, 3, "#2E5C8A"),    #  0 → 3
         ("二  分篇撰写与初稿", 3, 7, "#3E7CB1"),         #  3 → 10
         ("三  教学试用与反馈", 10, 3, "#5B9BD5"),        # 10 → 13
         ("四  统稿修改与鉴定", 13, 3, "#3E7CB1"),        # 13 → 16
         ("五  定稿出版与推广", 16, 6, "#2E5C8A")],       # 16 → 22
        {3: "大纲审定", 10: "初稿完成", 13: "试用启动",
         16: "提交出版", 22: "出版发行"},
        title="图3-1  编写工作进度甘特图（2026.09—2028.06）")

    merge_figures([f1, f2, f3], out_path)
    return out_path


def main():
    out = os.path.join(HERE, "demo.pptx")
    build(out)
    print("[1/3] 已生成:", out)

    # 几何自检
    issues = verify(out, verbose=True)
    if issues:
        print("[2/3] 几何自检发现问题:")
        for it in issues:
            print("   -", it)
    else:
        print("[2/3] 几何自检通过：无浮点坐标、无越界图元、箭头端点全部贴合")

    # 字符画预览（需要 Windows + PowerPoint）
    try:
        import check_render
        print("[3/3] 字符画预览：")
        check_render.report(out)
    except Exception as e:
        print("[3/3] 字符画预览跳过（%s）" % type(e).__name__)
        print("      需 Windows 且已安装 PowerPoint；绘图与自检不受影响。")


if __name__ == "__main__":
    main()
