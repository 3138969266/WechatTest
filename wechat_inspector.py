# -*- coding: utf-8 -*-
"""
微信UI结构诊断工具
用于查看微信PC客户端的UI自动化控件结构
帮助你了解如何定位微信消息
"""

import time
import sys

try:
    import uiautomation as auto
    UIAUTO_AVAILABLE = True
except ImportError:
    UIAUTO_AVAILABLE = False
    print("错误: 未安装uiautomation库")
    print("请运行: pip install uiautomation")
    sys.exit(1)

try:
    import win32gui
    import win32con
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("警告: 未安装pywin32库，部分功能受限")
    print("请运行: pip install pywin32")


def print_tree(control, indent=0, max_depth=5, show_all=False):
    """
    递归打印控件树结构
    
    Args:
        control: UI控件
        indent: 缩进级别
        max_depth: 最大深度
        show_all: 是否显示所有控件（包括空名称的）
    """
    if indent > max_depth:
        return
    
    prefix = "  " * indent
    
    try:
        control_type = control.ControlTypeName
        name = control.Name
        automation_id = control.AutomationId
        class_name = control.ClassName
        
        # 决定是否显示这个控件
        should_show = show_all
        if not should_show:
            # 显示有名称、有AutomationId或特定类型的控件
            if (name and name.strip()) or automation_id or control_type in [
                'WindowControl', 'PaneControl', 'ListControl', 'ListItemControl',
                'TextControl', 'ButtonControl', 'EditControl', 'TabControl'
            ]:
                should_show = True
        
        if should_show:
            # 构建显示信息
            info_parts = [control_type]
            
            if name and name.strip():
                # 截断过长的名称
                display_name = name[:50] + "..." if len(name) > 50 else name
                info_parts.append(f'Name="{display_name}"')
            
            if automation_id:
                info_parts.append(f'Id="{automation_id}"')
            
            if class_name and class_name not in ['', 'CtrlNotifySink', 'ATL:0000']:
                info_parts.append(f'Class="{class_name}"')
            
            print(f"{prefix}{' | '.join(info_parts)}")
        
        # 递归处理子控件
        try:
            children = control.GetChildren()
            for child in children:
                print_tree(child, indent + 1, max_depth, show_all)
        except Exception as e:
            if indent < 2:  # 只在较高级别显示错误
                print(f"{prefix}  [获取子控件出错: {str(e)}]")
                
    except Exception as e:
        print(f"{prefix}[控件访问出错: {str(e)}]")


def find_wechat_window(window_title="微信"):
    """查找微信窗口"""
    # 方法1: 使用uiautomation
    print(f"\n正在查找窗口标题为 '{window_title}' 的微信窗口...")
    
    try:
        wechat_win = auto.WindowControl(Name=window_title)
        if wechat_win.Exists(maxSearchSeconds=3):
            print(f"✓ 找到微信窗口 (uiautomation)")
            return wechat_win
    except Exception as e:
        print(f"  uiautomation查找出错: {str(e)}")
    
    # 方法2: 枚举所有窗口，查找包含"微信"的
    print("\n正在枚举所有顶级窗口，查找微信相关窗口...")
    
    windows = []
    
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                windows.append((hwnd, title))
        return True
    
    if WIN32_AVAILABLE:
        win32gui.EnumWindows(callback, None)
        
        # 显示包含"微信"或"WeChat"的窗口
        wechat_windows = []
        for hwnd, title in windows:
            if '微信' in title or 'WeChat' in title or 'wechat' in title:
                wechat_windows.append((hwnd, title))
                print(f"  找到窗口: HWND={hwnd}, Title='{title}'")
        
        if wechat_windows:
            # 尝试用第一个找到的窗口
            first_title = wechat_windows[0][1]
            print(f"\n尝试使用窗口标题: '{first_title}'")
            try:
                wechat_win = auto.WindowControl(Name=first_title)
                if wechat_win.Exists(maxSearchSeconds=2):
                    print(f"✓ 成功获取窗口控件")
                    return wechat_win
            except Exception as e:
                print(f"  获取控件出错: {str(e)}")
    else:
        print("  (pywin32未安装，无法枚举窗口)")
    
    return None


def analyze_specific_area(wechat_win, area_name="", search_depth=6):
    """
    分析微信窗口的特定区域
    
    Args:
        wechat_win: 微信窗口控件
        area_name: 区域名称（用于提示）
        search_depth: 搜索深度
    """
    print(f"\n{'='*60}")
    print(f"分析区域: {area_name}")
    print(f"{'='*60}")
    
    # 1. 先显示概要（不显示空名称控件）
    print("\n[控件结构概要] (仅显示有名称或重要控件)")
    print("-" * 60)
    print_tree(wechat_win, max_depth=search_depth, show_all=False)
    
    # 2. 询问是否显示详细信息
    try:
        response = input("\n是否显示完整控件结构（包括空名称控件）? (y/n): ").strip().lower()
        if response in ['y', 'yes', '是']:
            print("\n[完整控件结构]")
            print("-" * 60)
            print_tree(wechat_win, max_depth=search_depth, show_all=True)
    except:
        pass


def search_for_keywords(wechat_win, keywords):
    """
    在控件树中搜索包含特定关键词的控件
    
    Args:
        wechat_win: 微信窗口控件
        keywords: 关键词列表
    """
    print(f"\n{'='*60}")
    print(f"搜索关键词: {keywords}")
    print(f"{'='*60}")
    
    results = []
    
    def search(control, depth=0, max_depth=10):
        if depth > max_depth:
            return
        
        try:
            name = control.Name
            control_type = control.ControlTypeName
            automation_id = control.AutomationId
            
            # 检查名称是否包含关键词
            if name:
                for keyword in keywords:
                    if keyword.lower() in name.lower():
                        results.append({
                            'depth': depth,
                            'type': control_type,
                            'name': name,
                            'automation_id': automation_id
                        })
                        break
            
            # 递归搜索子控件
            try:
                for child in control.GetChildren():
                    search(child, depth + 1, max_depth)
            except:
                pass
                
        except:
            pass
    
    search(wechat_win)
    
    if results:
        print(f"\n找到 {len(results)} 个匹配的控件:\n")
        for i, r in enumerate(results, 1):
            indent = "  " * min(r['depth'], 3)
            info = f"{indent}{r['type']}"
            if r['name']:
                info += f' Name="{r["name"]}"'
            if r['automation_id']:
                info += f' Id="{r["automation_id"]}"'
            print(f"{i}. {info}")
    else:
        print("\n未找到匹配的控件")
    
    return results


def interactive_inspection():
    """交互式检查"""
    print("\n" + "="*60)
    print("微信UI自动化诊断工具")
    print("="*60)
    print("\n本工具将帮助你：")
    print("1. 查找微信窗口")
    print("2. 分析微信的UI控件结构")
    print("3. 搜索特定关键词（如'未读消息'、'消息'等）")
    print("\n使用前请确保：")
    print("✓ 微信PC客户端已登录")
    print("✓ 微信窗口已打开（不要最小化到托盘）")
    print("✓ 可以看到微信的主界面")
    
    input("\n按回车键开始...")
    
    # 查找微信窗口
    wechat_win = find_wechat_window()
    
    if not wechat_win:
        print("\n" + "!"*60)
        print("错误: 无法找到微信窗口")
        print("!"*60)
        print("\n可能的原因：")
        print("1. 微信未登录或未启动")
        print("2. 微信窗口最小化到了托盘")
        print("3. 微信窗口标题不是'微信'")
        print("\n请尝试：")
        print("- 点击任务栏微信图标，打开微信窗口")
        print("- 查看微信窗口的实际标题")
        return
    
    # 主循环
    while True:
        print("\n" + "="*60)
        print("请选择操作：")
        print("="*60)
        print("1. 分析整个微信窗口结构（推荐先执行这个）")
        print("2. 搜索'未读消息'相关控件")
        print("3. 搜索'消息'相关控件")
        print("4. 搜索自定义关键词")
        print("5. 查看文本控件列表（可能包含消息内容）")
        print("0. 退出")
        
        choice = input("\n请输入选项 (0-5): ").strip()
        
        if choice == '0':
            print("\n感谢使用！")
            break
            
        elif choice == '1':
            try:
                depth = int(input("请输入搜索深度 (建议4-8，默认6): ") or "6")
            except:
                depth = 6
            analyze_specific_area(wechat_win, "整个微信窗口", search_depth=depth)
            
        elif choice == '2':
            search_for_keywords(wechat_win, ['未读', '条消息', '未读消息'])
            
        elif choice == '3':
            search_for_keywords(wechat_win, ['消息', 'Message', 'Chat'])
            
        elif choice == '4':
            keyword = input("请输入要搜索的关键词: ").strip()
            if keyword:
                search_for_keywords(wechat_win, [keyword])
            else:
                print("未输入关键词")
            
        elif choice == '5':
            print("\n" + "="*60)
            print("查找所有文本控件...")
            print("="*60)
            
            text_controls = []
            
            def find_text(control, depth=0, max_depth=10):
                if depth > max_depth:
                    return
                try:
                    if control.ControlTypeName == 'TextControl':
                        name = control.Name
                        if name and name.strip():
                            text_controls.append({
                                'depth': depth,
                                'name': name
                            })
                    try:
                        for child in control.GetChildren():
                            find_text(child, depth + 1, max_depth)
                    except:
                        pass
                except:
                    pass
            
            find_text(wechat_win)
            
            if text_controls:
                print(f"\n找到 {len(text_controls)} 个文本控件:\n")
                for i, tc in enumerate(text_controls, 1):
                    indent = "  " * min(tc['depth'], 3)
                    # 截断过长的文本
                    display_name = tc['name'][:60] + "..." if len(tc['name']) > 60 else tc['name']
                    print(f"{i}. {indent}{repr(display_name)}")
            else:
                print("\n未找到文本控件")
        
        else:
            print("无效选项，请重新选择")


def quick_scan():
    """快速扫描模式"""
    print("\n" + "="*60)
    print("快速扫描模式")
    print("="*60)
    
    wechat_win = find_wechat_window()
    
    if not wechat_win:
        return
    
    print("\n正在执行快速扫描...")
    
    # 1. 搜索未读消息
    print("\n[1/3] 搜索未读消息提示...")
    unread_results = []
    
    def search_unread(control, depth=0):
        if depth > 6:
            return
        try:
            name = control.Name
            if name and ('未读' in name or '条消息' in name):
                unread_results.append((control.ControlTypeName, name))
            for child in control.GetChildren():
                search_unread(child, depth + 1)
        except:
            pass
    
    search_unread(wechat_win)
    
    if unread_results:
        print(f"  找到 {len(unread_results)} 个未读消息相关控件:")
        for ctrl_type, name in unread_results:
            print(f"    - {ctrl_type}: {name}")
    else:
        print("  未找到未读消息提示（可能当前没有未读消息）")
    
    # 2. 搜索文本控件
    print("\n[2/3] 搜索文本控件...")
    text_controls = []
    
    def search_text(control, depth=0):
        if depth > 8:
            return
        try:
            if control.ControlTypeName == 'TextControl':
                name = control.Name
                if name and name.strip():
                    text_controls.append(name)
            for child in control.GetChildren():
                search_text(child, depth + 1)
        except:
            pass
    
    search_text(wechat_win)
    
    if text_controls:
        print(f"  找到 {len(text_controls)} 个文本控件")
        print("  前10个:")
        for i, name in enumerate(text_controls[:10], 1):
            display = name[:40] + "..." if len(name) > 40 else name
            print(f"    {i}. {repr(display)}")
    else:
        print("  未找到文本控件")
    
    # 3. 显示控件结构概要
    print("\n[3/3] 显示控件结构概要...")
    print_tree(wechat_win, max_depth=4, show_all=False)
    
    print("\n" + "="*60)
    print("快速扫描完成！")
    print("="*60)
    print("\n提示：")
    print("- 如果你看到了联系人名称或消息内容，说明UI自动化可以获取到数据")
    print("- 运行 '交互式检查' 可以查看更详细的结构")


if __name__ == '__main__':
    print("="*60)
    print("微信UI自动化诊断工具")
    print("="*60)
    print("\n选择运行模式：")
    print("1. 快速扫描（自动执行常用检查）")
    print("2. 交互式检查（详细分析）")
    
    mode = input("\n请选择 (1/2，默认1): ").strip()
    
    if mode == '2':
        interactive_inspection()
    else:
        quick_scan()
