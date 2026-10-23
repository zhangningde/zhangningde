#!/usr/bin/env python3
"""Build operator manual V3 using the uploaded 20260727 template style."""

from __future__ import annotations

import shutil
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "source_manual" / "5G+IMS AI智能远程协助系统操作员使用手册20260727.docx"
ASSETS = ROOT / "manual_assets"
OUTPUT = ROOT / "output" / "5G+IMS_AI智能远程协助系统操作员使用手册_20261023_V3.0.docx"

TITLE_STYLE = "样式 样式 标题 1 + 右 + 左侧:  0 厘米 首行缩进:  0 厘米 段前: 15.6 磅"
BLUE = "1F4E79"


def clear_body(doc: Document) -> None:
    body = doc.element.body
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            continue
        body.remove(child)


def set_run(run, *, size=11, bold=None, color=None, font="宋体") -> None:
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def p(
    doc: Document,
    text: str,
    *,
    style="Normal",
    size=11,
    bold=False,
    align=None,
    indent=False,
    space_after=6,
    color=None,
):
    para = doc.add_paragraph(style=style)
    if align is not None:
        para.alignment = align
    if indent:
        para.paragraph_format.first_line_indent = Cm(0.74)
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = para.add_run(text)
    set_run(run, size=size, bold=bold, color=color)
    return para


def section_title(doc: Document, text: str) -> None:
    try:
        para = doc.add_paragraph(text, style=TITLE_STYLE)
    except KeyError:
        para = doc.add_paragraph(text, style="Heading 1")
    for run in para.runs:
        set_run(run, size=15, bold=True, color=BLUE, font="黑体")


def h(doc: Document, text: str, *, size=12) -> None:
    p(doc, text, size=size, bold=True, color=BLUE, space_after=8, font_override=True)


def p(doc, text, *, style="Normal", size=11, bold=False, align=None, indent=False, space_after=6, color=None, font_override=False):
    para = doc.add_paragraph(style=style)
    if align is not None:
        para.alignment = align
    if indent:
        para.paragraph_format.first_line_indent = Cm(0.74)
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = para.add_run(text)
    set_run(run, size=size, bold=bold, color=color, font="黑体" if font_override and bold else "宋体")
    return para


def ops(doc: Document, text: str) -> None:
    p(doc, f"操作说明：{text}", style="Body Text", indent=True, space_after=8)


def note(doc: Document, text: str) -> None:
    p(doc, f"注意：{text}", indent=True, bold=True, color="C00000", space_after=8)


def bullet(doc: Document, text: str) -> None:
    para = doc.add_paragraph(style="Body Text")
    para.paragraph_format.left_indent = Cm(0.74)
    para.paragraph_format.first_line_indent = Cm(-0.37)
    para.paragraph_format.space_after = Pt(3)
    set_run(para.add_run(f"•    {text}"), size=10.5)


def caption(doc: Document, text: str) -> None:
    p(doc, text, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, color="666666", space_after=10)


def picture(doc: Document, filename: str, caption_text: str, width_cm: float = 15.2) -> None:
    path = ASSETS / filename
    if not path.exists():
        p(doc, f"【缺少截图：{filename}】", align=WD_ALIGN_PARAGRAPH.CENTER, color="C00000")
        caption(doc, caption_text)
        return
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(4)
    para.paragraph_format.space_after = Pt(2)
    para.add_run().add_picture(str(path), width=Cm(width_cm))
    caption(doc, caption_text)


def update_header_footer(doc: Document) -> None:
    header = doc.sections[0].header
    if header.paragraphs:
        header.paragraphs[0].text = ""
        run = header.paragraphs[0].add_run("5G+IMS AI智能远程协助系统操作员手册")
        set_run(run, size=9, color="666666")
        header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer = doc.sections[0].footer
    if footer.paragraphs:
        footer.paragraphs[0].text = ""
        run = footer.paragraphs[0].add_run("版本：3.0（2026-10-23改版适配）                                                                                               ")
        set_run(run, size=9)


def cover(doc: Document) -> None:
    for _ in range(5):
        doc.add_paragraph()
    p(doc, "US2000系列", size=26, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, color=BLUE, space_after=4)
    p(doc, "5G+IMS AI智能远程协助系统", size=26, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, color=BLUE, space_after=4)
    p(doc, "操作员手册", size=26, bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, color=BLUE, space_after=18)
    p(doc, "改版功能适配版  V3.0", size=14, align=WD_ALIGN_PARAGRAPH.RIGHT, color="666666", space_after=40)
    p(doc, "广东中为通信技术有限公司", size=15, bold=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=20)
    doc.add_page_break()


def disclaimer(doc: Document) -> None:
    p(
        doc,
        "本手册内容如有变化，恕不另行通知，手册的内容也不作为广东中为通信技术有限公司的承诺。本公司将保留根据工程和生产技术的发展对产品设计进行修改的权利，不另行通知。版权所有，广东中为通信技术有限公司印制。",
        style="Body Text",
        size=11,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        space_after=16,
    )


def front(doc: Document) -> None:
    section_title(doc, "文件概述")
    p(
        doc,
        "本文档内容涉及US2000系列产品AI智能远程协助系统界面功能介绍、应用指导说明和使用方法，适用于用户操作人员。本版以《操作员使用手册20260727》为模版风格，结合2026年10月23日现网Web操作台改版后的实际功能与界面重新编制。",
        indent=True,
    )

    section_title(doc, "敬请留意")
    p(doc, "在开始使用前，请仔细阅读本节内容。", indent=True)
    p(doc, "安装须知", size=18, bold=True)
    p(
        doc,
        "本系统支持AlmaLinux、RockyLinux、OracleLinux、RedHat、Debian、Ubuntu、CentOS、龙蜥、欧拉、麒麟操作系统。",
        indent=True,
    )
    p(doc, "注意事项", size=18, bold=True)
    p(doc, "因为文档中涉及到用户的账号信息，请务必确保用户信息安全。", indent=True)
    p(doc, "确保自己的登录账号不被泄露。", indent=True)
    p(doc, "本手册不写入生产密码；现场密码由管理员单独发放。", indent=True)

    section_title(doc, "前言")
    p(doc, "欢迎选购", bold=True)
    p(doc, "感谢使用5G+IMS AI智能远程协助系统!", style="Body Text", indent=True)
    p(doc, "读者对象", bold=True)
    bullet(doc, "相关业务操作人员")
    p(doc, "如何使用本手册", bold=True)
    p(
        doc,
        "本文档内容为5G+IMS AI智能远程协助系统用户业务功能使用手册。在开始使用之前，请仔细阅读本手册各章节内容。",
        style="Body Text",
        indent=True,
    )
    p(doc, "使用须知", bold=True)
    bullet(
        doc,
        "系统采用B/S架构，用户可以通过Windows、Linux、安卓、欧拉、麒麟、苹果等操作系统访问系统管理页面。",
    )
    p(doc, "注意事项", bold=True)
    bullet(
        doc,
        "本系统基于ARM和X86、C86架构和国产化的硬件系统硬件架构，单套系统在硬件配置满足的情况下，对用户的业务容量和交换能力无限制。",
    )
    bullet(doc, "调度坐席支持软终端注册接入和硬件终端接入两种方式。")
    bullet(doc, "本手册以操作员账号在Web操作台可见功能为准；不同权限、不同部署或后续升级可能导致局部界面差异。")
    doc.add_page_break()


def chapter_1(doc: Document) -> None:
    section_title(doc, "1、产品简介")
    p(doc, "1.1 概述", style="Body Text", size=12, bold=True, color=BLUE)
    p(
        doc,
        "UCS2000系列5G+IMS AI智能远程协助系统是一款基于ARM、X86、C86国产化硬件架构灵活部署的超大容量、高可靠性、安全要求的智能化远程协助系统。系统支持国产化操作系统，支持IPv4和IPv6全栈兼容能力，基于模块化设计，可根据实际接入用户数量灵活选择不同的硬件配置，快速搭建专用通信平台。",
        style="Body Text",
        indent=True,
    )
    p(
        doc,
        "系统集融合调度指挥、远程协助、呼叫中心、应急通信、数字程控交换、IMS、Centrex、CRM、录音存证、视频调度、会议会商、对讲融合、广播、应急通知等多种应用服务于一体，旨在为大中小企业、军队、石油、电力、运营商以及专用行业用户提供高效、可靠的高速融合统一综合性通信能力。",
        style="Body Text",
        indent=True,
    )
    p(doc, "1.2 关键特性", style="Body Text", size=12, bold=True, color=BLUE)
    for item in [
        "系统基于ARM、X86、C86国产统信处理器架构，硬件可根据用户业务模块以及接入用户数需求选择配置，部署方便快捷。",
        "扩容方便，只需要扩容硬件，就可以快速扩容。",
        "系统功能丰富，基于模块化设计，可为用户提供融合调度指挥、远程协助、呼叫中心、应急通信、数字程控交换、IMS、Centrex、CRM、录音存证、视频调度、会议会商、对讲融合、广播、应急通知等多种应用服务。",
        "系统容量大，单套系统在硬件满足要求情况下，可支持10万坐席以及2万用户的并发能力【需要硬件服务器支撑】。",
        "系统自带安全加密、安全审计功能，能实现安全可靠的认证加密对接，支持与核心交换设备端到端加密通信。",
        "支持自主防火墙，系统支持安全访问控制、非法地址自动拦截、恶意注册自动拦截、恶意呼叫自动屏蔽、安全反诈甄别等安全保障能力。",
        "系统具备与无线专网应急通信系统对接能力，支持基于Web统一实时监控、维护网络管理功能。",
        "操作员Web操作台当前可见业务模块包括：语音调度、视频会议、会议会商、录音录像。",
    ]:
        bullet(doc, item)


def chapter_2(doc: Document) -> None:
    section_title(doc, "2、系统登录")
    p(doc, "2.1 用户介绍", style="Body Text", size=12, bold=True, color=BLUE)
    p(doc, "本系统分为两种用户类型：", indent=True)
    p(
        doc,
        "企业管理员(admin)：作为系统初始化管理帐号，负责系统日常维护、业务数据管理、系统安全管理能力。管理员对企业有最高权限。",
        style="Body Text",
        indent=True,
    )
    p(
        doc,
        "操作员：帐号由管理员设定后发放。普通操作员座席的可用功能受限于组织机构的权限体系，在权限允许的情况下，使用可用的功能。",
        indent=True,
    )

    p(doc, "2.2 系统登录", style="Body Text", size=12, bold=True, color=BLUE)
    ops(
        doc,
        "登录系统时，推荐使用Google浏览器或Microsoft Edge。地址栏输入系统入口，如 https://36.141.81.170:8091/webuser ，点击回车进入登录界面。在【请输入用户名】和【请输入密码】输入框中填写已分配的账号信息，确认无误后单击【登录】。",
    )
    picture(doc, "010-系统登录页.png", "图 2.2  系统登录")
    p(doc, "系统登录成功后，可以进入到业务功能页面。请确认顶部显示【状态服务已连接】，分机状态为在线空闲，并检查麦克风、扬声器和摄像头是否正常。", indent=True)

    p(doc, "2.3 系统登出", style="Body Text", size=12, bold=True, color=BLUE)
    ops(
        doc,
        "系统退出登录时，点击右上角退出图标，即可从系统中注销登录。注销登录之后系统页面自动退回到登录页面。公共值班终端交接时，不得仅关闭标签页或让浏览器保存密码。",
    )


def chapter_3(doc: Document) -> None:
    section_title(doc, "3、语音调度与视频会议")
    p(doc, "3.1 语音调度", size=12, bold=True, color=BLUE)
    ops(
        doc,
        "在顶部功能导航选择【语音调度】，可通过拨号键盘或通讯录发起音频、视频呼叫。页面同时显示快捷联系人、当前通话和最近通话。操作前请确认坐席显示为在线，并检查麦克风、扬声器和摄像头是否正常。",
    )
    picture(doc, "020-语音调度工作台.png", "图 3.1  语音调度")

    p(doc, "3.1.1 通话调度拨号盘", bold=True, color=BLUE)
    ops(
        doc,
        "在【被叫号码】中直接输入号码，或使用数字拨号盘完成输入；单击语音呼叫发起音频呼叫，单击视频呼叫发起视频呼叫。删除键用于清除末位号码。未选择联系人或未输入有效号码时，呼叫按钮保持灰色不可用。",
    )
    picture(doc, "021-语音调度拨号键盘.png", "图 3.1.1  通话调度拨号盘")

    p(doc, "3.1.2 通讯录", bold=True, color=BLUE)
    p(doc, "3.1.2.1 Web通讯录", bold=True)
    ops(
        doc,
        "左侧【web通讯录】按部门组织显示联系人数量。可展开部门树查找目标人员，也可在【搜索部门、姓名或号码】中输入关键字定位。选中联系人后，根据业务需要发起音频或视频呼叫。",
    )
    p(doc, "3.1.2.2 快捷联系人与通话记录", bold=True)
    ops(
        doc,
        "【快捷联系人】用于展示常用对象；【当前通话】显示正在进行的会话；【最近通话】保存历史记录。可刷新通话记录，选中有效记录后按权限回拨。回拨前仍需确认对象和班次。",
    )

    p(doc, "3.2 视频会议", size=12, bold=True, color=BLUE)
    ops(
        doc,
        "切换至【视频会议】页签，通过通讯录选择参会人员并发起会议。会议建立后，可在成员区和底部工具栏执行成员管理及会场控制。未建立会议时，邀请、布局、成员、禁言和结束等按钮可能保持灰色。",
    )
    picture(doc, "030-视频会议工作台.png", "图 3.2  视频会议")

    p(doc, "3.2.1 视频会议通讯录", bold=True, color=BLUE)
    ops(
        doc,
        "通讯录支持按部门或分机名称搜索，也可展开组织树逐级查找。部门名称右侧显示当前条目数量。可使用通讯录设置按钮调整部门显示层级。",
    )
    p(doc, "3.2.2 视频会议成员操作", bold=True, color=BLUE)
    ops(
        doc,
        "在会议成员区域选择目标成员，可执行与该成员相关的呼叫、静音、视频或移除等操作。具体可用按钮以当前账号权限和成员状态为准。",
    )
    p(doc, "3.2.3 会议成员邀请", bold=True, color=BLUE)
    ops(
        doc,
        "单击【邀请】，在通讯录中选择需要加入会议的成员并确认。系统将向所选成员发起入会呼叫。邀请前请确认对象身份和会议敏感等级。",
    )
    p(doc, "3.2.4 会议操作", bold=True, color=BLUE)
    ops(
        doc,
        "会议底部工具栏提供邀请、布局、成员、禁言、更多和结束等功能。【更多】菜单中可进行带宽、帧率、位置、主屏、比例及会议锁定等设置。结束会议前请确认业务已办理完毕。",
    )
    picture(doc, "031-视频会议更多控制.png", "图 3.2.4  会议操作")

    for title, text in [
        ("3.2.4.1 会议成员邀请", "单击【邀请】，在通讯录中选择需要加入会议的成员并确认。系统将向所选成员发起入会呼叫。"),
        ("3.2.4.2 成员界面布局", "单击【布局】调整会议画面布局。人数变化后应检查主画面是否符合会商需要。"),
        ("3.2.4.3 成员列表", "单击【成员】查看当前参会成员，核对异常掉线、重复成员或未入会对象。"),
        ("3.2.4.4 发言和禁言", "单击【禁言】可执行全体禁言控制。发布重要指令前应说明禁言原因，并保留必要发言通道。"),
        ("3.2.4.5 带宽设置", "根据网络条件和会议质量要求选择合适的会议带宽。网络不稳定时可适当降低带宽，以提高会议连续性。"),
        ("3.2.4.6 帧率设置", "设置视频帧率。较高帧率画面更流畅，但会占用更多网络资源；应结合终端能力和网络状况选择。"),
        ("3.2.4.7 成员位置设置", "调整选定成员在会议画面布局中的显示位置，确认后观察布局是否符合会商需要。"),
        ("3.2.4.8 指定成员主屏设置", "选择目标成员并将其指定为主屏，使该成员画面在主显示区域呈现。"),
        ("3.2.4.9 屏幕显示比例设置", "根据显示设备和画面内容选择适合的显示比例，避免画面拉伸、裁切或出现过多留白。"),
        ("3.2.4.10 锁定和解锁会议设置", "锁定会议后可限制未获授权的成员加入；需要继续邀请成员时先解锁。操作后应核对会议锁定状态。"),
    ]:
        p(doc, title, bold=True, color=BLUE)
        ops(doc, text)


def chapter_4(doc: Document) -> None:
    section_title(doc, "4 会议会商")
    ops(
        doc,
        "在顶部功能导航选择【会议会商】，进入会议室列表及会商管理页面，可完成会议室维护、会议启动和会中控制。页面包含【会议列表】和【视频窗口】两个页签。",
    )
    picture(doc, "040-会议会商列表.png", "图 4  会议会商")

    p(doc, "4.1 会议室列表管理", size=12, bold=True, color=BLUE)
    ops(
        doc,
        "会议室列表展示会议室号码、备注、密码、开启视频、录音录像及操作按钮。可对会议室执行查询、添加、修改、删除、查看详情和开始会议。会议室是长期配置对象，不应随意创建或删除。",
    )

    p(doc, "4.1.1 会议室信息查询", bold=True, color=BLUE)
    ops(
        doc,
        "在查询区域输入会议室号码、会议室备注，单击【查询】获取匹配结果；单击【重置】清空条件并恢复默认列表。",
    )

    p(doc, "4.1.2 会议室管理", bold=True, color=BLUE)
    p(doc, "4.1.2.1 会议室添加", bold=True)
    ops(
        doc,
        "单击【添加】打开会议室配置窗口，依次完成基本设置、视频设置、录音录像及会议成员配置，检查必填项后提交。",
    )
    picture(doc, "042-添加会议-01-基本设置.png", "图 4.1.2.1  会议室添加", width_cm=13.2)

    p(doc, "4.1.2.1.1 基本信息", bold=True)
    ops(
        doc,
        "填写会议室号码、备注和密码，按需启用进入时播放、退出时播放、一人时播放、等待时播放提示音，并设置最大会议时长。界面提示会议室号码和备注最长12个字符，密码最长10个字符。带红色星号的项目为必填项。",
    )
    picture(doc, "042-添加会议-01-基本设置.png", "图 4.1.2.1.1  基本信息", width_cm=13.2)

    p(doc, "4.1.2.1.2 视频设置", bold=True)
    ops(
        doc,
        "配置会议默认视频参数，包括开启视频、视频分辨率、视频布局、视频质量、视频帧率、视频带宽、最大成员数量和显示成员名称。建议使用与终端和网络能力匹配的参数。",
    )
    picture(doc, "042-添加会议-02-视频设置.png", "图 4.1.2.1.2  视频设置", width_cm=13.2)

    p(doc, "4.1.2.1.3 录音设置", bold=True)
    ops(
        doc,
        "在录音录像页签中设置是否开启录音录像。启用后会议可形成MP4等媒体文件。涉及录音录像时，应遵守本单位授权、告知、存储和保密规定。",
    )
    picture(doc, "042-添加会议-03-录音录像.png", "图 4.1.2.1.3  录音设置", width_cm=13.2)

    p(doc, "4.1.2.1.4 成员管理", bold=True)
    ops(
        doc,
        "在成员管理页签中，从可选成员区域按组织树或搜索条件选择默认参会成员，核对已选成员名单后提交。只有明确需要重新选择时才使用【清空】。",
    )
    picture(doc, "042-添加会议-04-会议成员.png", "图 4.1.2.1.4  成员管理", width_cm=13.2)

    p(doc, "4.1.2.2 会议室修改", bold=True)
    ops(
        doc,
        "在会议室列表中选择目标会议室并单击【修改】，按需调整各页签配置。保存前核对会议室号码、媒体参数及成员信息，避免影响正在进行的会议。",
    )
    p(doc, "4.1.2.2.1 基本信息", bold=True)
    ops(doc, "修改会议室号码、备注、密码、提示音和最大会议时长。变更前应记录原参数和变更原因。")
    p(doc, "4.1.2.2.2 修改视频设置", bold=True)
    ops(doc, "进入视频设置页签修改会议画面及媒体参数，保存后在后续会议中验证显示效果和网络占用情况。")
    p(doc, "4.1.2.2.3 修改录音录像设置", bold=True)
    ops(doc, "根据管理要求启用或调整会议录音、录像选项。涉及录音录像时，应遵守本单位授权、告知、存储和保密规定。")
    p(doc, "4.1.2.2.4 修改预设会议成员", bold=True)
    ops(doc, "调整已选成员清单，增加或移除默认参会对象。提交前再次核对姓名、分机号和部门。")

    p(doc, "4.1.2.3 删除会议室", bold=True)
    ops(
        doc,
        "选择目标会议室并单击【删除】。出现确认对话框后，再次核对会议室号码，确认无业务依赖且已获授权再执行。",
    )
    note(doc, "执行删除操作前，请确认所选对象及数据留存要求。删除完成后应检查列表，确认操作结果。")

    p(doc, "4.1.2.4 会议室查看", bold=True)
    ops(
        doc,
        "选择会议室并单击【查看详情】，以只读方式核对基本信息、视频参数、录音录像设置和成员配置。",
    )
    picture(doc, "041-会议详情-01-基本设置.png", "图 4.1.2.4.1  基本信息", width_cm=13.2)
    picture(doc, "041-会议详情-02-视频设置.png", "图 4.1.2.4.2  视频设置", width_cm=13.2)
    picture(doc, "041-会议详情-03-录音录像.png", "图 4.1.2.4.3  录音设置", width_cm=13.2)
    picture(doc, "041-会议详情-04-会议成员.png", "图 4.1.2.4.4  成员管理", width_cm=13.2)

    p(doc, "4.1.2.5 开始会议", bold=True)
    ops(
        doc,
        "在会议室列表中选择目标会议室并单击【开始会议】。会议建立后，切换到【视频窗口】观察成员呼叫及入会状态，并通过底部工具栏进行会中控制。",
    )
    picture(doc, "043-会议会商视频窗口.png", "图 4.1.2.5  开始会议")

    for title, text in [
        ("4.1.2.5.1 会议成员邀请", "单击【邀请】，在通讯录中选择需要加入会议的成员并确认。系统将向所选成员发起入会呼叫。"),
        ("4.1.2.5.2 成员界面布局", "单击【布局】调整会议画面布局，确认主会场与现场画面位置合理。"),
        ("4.1.2.5.3 成员列表", "单击【成员】查看当前参会成员，核对入会状态。"),
        ("4.1.2.5.4 发言和禁言", "按会商需要执行禁言或恢复发言。重要指令发布前应说明控制意图。"),
        ("4.1.2.5.5 带宽设置", "根据网络条件和会议质量要求选择合适的会议带宽。网络不稳定时可适当降低带宽。"),
        ("4.1.2.5.6 帧率设置", "设置视频帧率。较高帧率画面更流畅，但会占用更多网络资源。"),
        ("4.1.2.5.7 成员位置设置", "调整选定成员在会议画面布局中的显示位置。"),
        ("4.1.2.5.8 指定成员主屏设置", "选择目标成员并将其指定为主屏，使该成员画面在主显示区域呈现。"),
        ("4.1.2.5.9 屏幕显示比例设置", "根据显示设备和画面内容选择适合的显示比例。"),
        ("4.1.2.5.10 锁定和解锁会议设置", "锁定会议后可限制未获授权的成员加入；需要继续邀请成员时先解锁。"),
    ]:
        p(doc, title, bold=True)
        ops(doc, text)


def chapter_5(doc: Document) -> None:
    section_title(doc, "5 录音录像查看")
    ops(
        doc,
        "在顶部功能导航选择【录音录像】，进入历史媒体文件管理页面。可按条件查询、播放查看或按权限删除记录。列表字段包括类型、业务类型、主叫号码、主叫名称、被叫/会议号、对应名称、开始时间、结束时间、录音时长、通话状态、录音录像文件和操作。",
    )
    picture(doc, "050-录音录像查询页.png", "图 5  录音录像查看")

    p(doc, "5.1 录音录像查询", size=12, bold=True, color=BLUE)
    ops(
        doc,
        "可按主叫号码、被叫号码或会议号、录音日期及类型组合筛选，单击【查询】显示结果；单击【重置】清空筛选条件。跨日业务应分别查询相关日期。",
    )
    picture(doc, "051-录音录像类型筛选.png", "图 5.1  录音录像查询")
    picture(doc, "052-录音录像日期筛选.png", "图 5.1.1  录音日期筛选")

    p(doc, "5.2 录音录像批量删除", size=12, bold=True, color=BLUE)
    ops(
        doc,
        "勾选需要删除的记录后单击【删除所选】，或在单条记录操作列中单击删除图标。系统弹出确认对话框后，再次核对所选对象再确认。普通值班操作员在没有书面授权时不得删除录音录像。",
    )
    note(
        doc,
        "执行删除操作前，请确认所选对象及数据留存要求。删除完成后应检查列表，确认操作结果。不得为了释放空间、隐藏操作失误或方便演示而删除生产录音录像。",
    )

    p(doc, "5.3 录音录像内容查看", size=12, bold=True, color=BLUE)
    ops(
        doc,
        "在结果列表中单击播放按钮，系统加载对应录音录像并弹出播放器展示。没有文件名的记录，播放按钮保持不可用。若提示无有效媒体，表示该记录未生成可播放文件。播放时应在受控环境进行，不得擅自复制或外传。",
    )
    p(
        doc,
        "通话状态说明：【正常通话挂机】表示会话按正常流程结束；【分机不存在】表示被叫号码未能匹配有效分机；【媒体超时】表示媒体链路未在规定时间内保持正常。状态异常时应结合录音时长、文件名和现场反馈判断。",
        indent=True,
    )


def chapter_6(doc: Document) -> None:
    section_title(doc, "6、改版功能补充说明")
    p(doc, "6.1 状态服务与分机状态", size=12, bold=True, color=BLUE)
    ops(
        doc,
        "登录后顶部显示【状态服务已连接】表示Web操作台与后台状态通道已建立；账号旁显示【在线空闲】表示当前分机已注册且未占用。若显示【状态服务未连接】，应先刷新页面并等待恢复，仍异常时保存截图并联系运维。",
    )
    p(doc, "6.2 AI能力说明", size=12, bold=True, color=BLUE)
    p(
        doc,
        "系统名称包含AI智能远程协助能力。本次操作员账号可见导航为语音调度、视频会议、会议会商、录音录像，页面中未出现独立的【AI助手】【AI降噪】【智能报告】【安全审计】菜单。相关能力可能由平台后台、媒体链路、移动终端或特定项目授权自动提供。操作员应以实际出现的按钮为准；若后续启用新的AI入口，应由管理员补充专项操作说明。",
        indent=True,
    )
    p(doc, "6.3 安全要求", size=12, bold=True, color=BLUE)
    bullet(doc, "实行一人一号，禁止共享账号或在群聊中发送密码。")
    bullet(doc, "离岗前退出系统；短时离开也应锁定操作终端。")
    bullet(doc, "通讯录、通话记录和录音录像可能包含个人信息与生产现场信息，应按最小必要原则使用。")
    bullet(doc, "删除会议室、结束会议、全体禁言、删除录音录像等高风险操作，必须确认授权和业务影响。")


def appendix(doc: Document) -> None:
    section_title(doc, "附录：常见问题与交接检查")
    p(doc, "A.1 常见问题", size=12, bold=True, color=BLUE)
    for q, a in [
        ("登录页打不开", "检查网络、白名单、VPN和8091端口，确认地址正确。"),
        ("登录失败", "重新核对账号密码各一次，禁止连续猜测；仍失败联系管理员。"),
        ("呼叫按钮灰色", "先选择通讯录联系人或输入有效被叫号码，并确认分机在线。"),
        ("呼叫无声或无画面", "检查耳麦、摄像头、浏览器权限及设备是否被其他软件占用。"),
        ("会议无法启动", "查看会议室配置、成员和权限，不重复点击开始会议。"),
        ("录音没有文件", "核对录制是否开启、通话状态、录音时长和存储服务状态。"),
    ]:
        p(doc, f"{q}：{a}", indent=True, space_after=4)

    p(doc, "A.2 上岗检查", size=12, bold=True, color=BLUE)
    for item in [
        "登录账号与当前值班人员一致。",
        "状态服务已连接，分机在线空闲。",
        "耳麦、麦克风和摄像头工作正常。",
        "通讯录能够加载和搜索。",
        "常用应急会议室配置存在。",
        "录音录像查询服务可用。",
    ]:
        bullet(doc, item)

    p(doc, "A.3 版本说明", size=12, bold=True, color=BLUE)
    p(
        doc,
        "本手册V3.0参照《5G+IMS AI智能远程协助系统操作员使用手册20260727》模版风格编制，并将现网改版后的操作内容纳入。主要更新包括：登录字段改为用户名/密码表述；导航名称对齐现网【语音调度】【视频会议】【会议会商】【录音录像】；补充状态服务、通话记录、会议室字段、录音状态释义及AI能力边界；配图替换为2026年10月23日现网脱敏截图。",
        indent=True,
    )


def build() -> Path:
    if not SOURCE.exists():
        raise FileNotFoundError(SOURCE)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(SOURCE, OUTPUT)
    doc = Document(OUTPUT)
    clear_body(doc)
    update_header_footer(doc)
    cover(doc)
    disclaimer(doc)
    front(doc)
    chapter_1(doc)
    chapter_2(doc)
    chapter_3(doc)
    chapter_4(doc)
    chapter_5(doc)
    chapter_6(doc)
    appendix(doc)
    doc.core_properties.title = "5G+IMS AI智能远程协助系统操作员使用手册"
    doc.core_properties.subject = "参照20260727模版风格的改版适配手册"
    doc.core_properties.author = "广东中为通信技术有限公司"
    doc.core_properties.comments = "V3.0 模版风格 + 现网改版内容"
    doc.save(OUTPUT)
    return OUTPUT


if __name__ == "__main__":
    # Fix duplicate p definition: keep the later one with font_override
    print(build())
