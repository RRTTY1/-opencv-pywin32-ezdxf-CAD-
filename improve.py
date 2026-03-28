import win32com.client
import time
import pythoncom
from pathlib import Path
import win32gui
import win32api
import win32con


def force_cad_active(doc, acad, last_mouse_pos=None):
    try:
        cad_hwnd = acad.HWND
        if not cad_hwnd:
            return last_mouse_pos

        if last_mouse_pos is None:
            last_mouse_pos = win32api.GetCursorPos()

        # 真的移动1像素鼠标，再移回来，强制CAD保持活跃
        current_x, current_y = win32api.GetCursorPos()
        win32api.SetCursorPos((current_x + 1, current_y))
        time.sleep(0.001)
        win32api.SetCursorPos((current_x, current_y))

        # 高频发送鼠标移动消息，双保险
        left, top, right, bottom = win32gui.GetClientRect(cad_hwnd)
        import random
        fake_x = left + random.randint(50, 200)
        fake_y = top + random.randint(50, 200)
        lparam = win32api.MAKELONG(fake_x, fake_y)
        for _ in range(3):
            win32gui.PostMessage(cad_hwnd, win32con.WM_MOUSEMOVE, 0, lparam)

        pythoncom.PumpWaitingMessages()
    except Exception as e:
        print(f"活跃维持异常: {e}")
        pass

    return last_mouse_pos


def connect_autocad():
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    try:
        acad = win32com.client.dynamic.Dispatch("AutoCAD.Application.25")
        print("已连接到已运行的AutoCAD")
    except Exception as e:
        print(f"连接已运行实例失败: {e}，尝试新建实例")
        try:
            acad = win32com.client.Dispatch("AutoCAD.Application.25")
            acad.Visible = True
            print("已启动AutoCAD，等待软件就绪...")
            time.sleep(10)
        except Exception as e2:
            print(f"新建实例失败: {e2}")
            pythoncom.CoUninitialize()
            return None, None, None

    try:
        if acad.Documents.Count == 0:
            acad.Documents.Add()
            print("已新建空白CAD文档")
            time.sleep(3)
        doc = acad.ActiveDocument
        doc.SendCommand("_ESC _ESC ")
        time.sleep(1)
        space = doc.ModelSpace
        print("成功获取模型空间，COM接口可用")
        # 【修正1】返回顺序和主程序解包顺序严格一致
        return acad, doc, space
    except Exception as e:
        print(f"获取文档/模型空间失败: {e}")
        pythoncom.CoUninitialize()
        return None, None, None


# ------------------- 主程序执行 -------------------
locap = Path(r"E:\Python\project\cv2\location\contour_coords.txt")
if __name__ == "__main__":
    # 连接CAD，获取应用对象和模型空间对象
    connect_result = connect_autocad()
    # 【修正2】解包顺序和函数返回顺序完全一致
    acad, doc, space = connect_result
    # 强制校验CAD连接有效性，避免无效对象报错
    if not acad or not doc or not space:
        print("CAD连接失败，程序退出")
        exit()

    try:
        # ------------------- 读取逻辑完全不变 -------------------
        contour_dict = {}
        current_contour_id = None
        current_points = []

        with open(locap, "r", encoding="utf-8") as f:
            for line in f:
                line_stripped = line.strip()
                if line_stripped.startswith("=== 帧") and line_stripped.endswith("==="):
                    if current_contour_id is not None and current_points and len(current_points) >= 3:
                        contour_dict[current_contour_id].append(current_points)
                    current_contour_id = line_stripped.replace("=== 帧", "").replace("===", "").strip()
                    contour_dict[current_contour_id] = []
                    current_points = []
                    continue
                elif not line_stripped:
                    if current_contour_id is not None and current_points and len(current_points) >= 3:
                        contour_dict[current_contour_id].append(current_points)
                    current_points = []
                    continue
                else:
                    try:
                        x, y = map(float, line_stripped.split(","))
                        # 过滤异常坐标（根据你的视频分辨率调整，比如1920x1080）
                        if 10 < x < 1920 and 10 < y < 1080:
                            current_points.append((x, y))
                    except:
                        print(f"跳过格式错误的行：{line_stripped}")
                        continue

            # 最后一帧的保存逻辑
            if current_contour_id is not None and current_points and len(current_points) >= 3:
                contour_dict[current_contour_id].append(current_points)

        print(f"✅ 成功读取 {len(contour_dict)} 帧数据")
        if not contour_dict:
            print("❌ 未读取到有效数据，程序退出")
            exit()

        sorted_frames = sorted(contour_dict.keys(), key=lambda x: int(x))
        total_frames = len(sorted_frames)

        # ------------------- 【修正3】ScreenUpdating属性访问，对象已正确，不会报错 -------------------
        # 先校验属性是否可访问，兼容破解版CAD
        try:
            original_screenupdating = acad.ScreenUpdating
            use_screen_control = True
        except Exception as e:
            print(f"⚠️  屏幕更新控制不可用，跳过: {e}")
            original_screenupdating = None
            use_screen_control = False

        last_mouse_pos = None
        try:
            if use_screen_control:
                acad.ScreenUpdating = False

            for id, contour_id in enumerate(sorted_frames, 1):
                last_mouse_pos = force_cad_active(doc, acad, last_mouse_pos)
                sub_contours = contour_dict[contour_id]
                frames = []

                try:
                    # 遍历当前帧的每个子轮廓，单独绘制
                    for points in sub_contours:
                        # 过滤点数不足的轮廓
                        if len(points) < 3:
                            continue

                        point = []
                        for (x, y) in points:
                            point.extend([x, y, 0.0])
                        points_variant = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, point)
                        # 绘制多段线
                        polyline = space.AddPolyline(points_variant)
                        # 设置线条属性
                        polyline.Color = 7
                        polyline.Lineweight = 50

                        first_point = points[0]
                        last_point = points[-1]
                        # 计算首尾点的距离，小于1.5像素才闭合
                        distance = ((first_point[0] - last_point[0]) ** 2 + (
                                    first_point[1] - last_point[1]) ** 2) ** 0.5
                        if distance < 1.5:
                            polyline.Closed = True
                        frames.append(polyline)

                        # 每画一个子轮廓也强制活跃
                        last_mouse_pos = force_cad_active(doc, acad, last_mouse_pos)

                    # 显示当前帧
                    if use_screen_control:
                        acad.ScreenUpdating = True
                    # 只在第一帧缩放一次，后续用轻量刷新
                    if id == 1:
                        acad.ZoomExtents()
                    else:
                        doc.SendCommand("_REDRAW ")

                    print(f"正在播放第 {id}/{total_frames} 帧", end="\r")
                    #time.sleep(0.08)  # 帧停留时间，让人眼能看到

                    # 关闭屏幕更新，准备后台删除
                    if use_screen_control:
                        acad.ScreenUpdating = False

                except Exception as e:
                    print(f"绘图失败: {e}")

                # 每帧结束后强制活跃
                last_mouse_pos = force_cad_active(doc, acad, last_mouse_pos)

                # 后台删除当前帧
                try:
                    for frame in frames:
                        frame.Delete()
                    doc.SendCommand("_REDRAW ")
                except Exception as del_e:
                    print(f"删除失败: {del_e}")
                    continue

            print("\n🎉 所有帧播放完成！")

        finally:
            # 【必须】恢复CAD原设置，避免卡死
            if use_screen_control and original_screenupdating is not None:
                acad.ScreenUpdating = original_screenupdating
    finally:
        # 释放COM环境
        pythoncom.CoUninitialize()