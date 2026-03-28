import win32com.client
import time
import pythoncom
from pathlib import Path

def force_cad_active(doc):
    #强制CAD保持活跃
    try:

        doc.SendCommand(" ")
        pythoncom.PumpWaitingMessages()
    except:
        pass


def connect_autocad():
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    acad = None
    space = None
    try:
        acad = win32com.client.dynamic.Dispatch("AutoCAD.Application.25")
        print("已连接到已运行的AutoCAD")
    except Exception as e:
        print(f"连接已运行实例失败: {e}，尝试新建实例")
        try:
            # 新建CAD实例
            acad = win32com.client.Dispatch("AutoCAD.Application.25")
            acad.Visible = True
            print("已启动AutoCAD，等待软件就绪...")
            time.sleep(10)
        except Exception as e2:
            print(f"新建实例失败: {e2}")
            pythoncom.CoUninitialize()
            return None

    try:
        # 校验并激活的文档，无文档则新建空白文档
        if acad.Documents.Count == 0:
            acad.Documents.Add()
            print("已新建空白CAD文档")
            time.sleep(3)
        doc = acad.ActiveDocument
        doc.SendCommand("_ESC _ESC ")
        time.sleep(1)
        # 获取模型空间
        space = doc.ModelSpace
        print("成功获取模型空间，COM接口可用")
        return acad, space, doc
    except Exception as e:
        print(f"获取文档/模型空间失败: {e}")
        pythoncom.CoUninitialize()
        return None , None, None


# --------------------------------------
locap=Path(r"E:\Python\project\cv2\location\contour_coords.txt")
if __name__ == "__main__":
    # 连接CAD
    connect_result = connect_autocad()
    if not connect_result  or len(connect_result) < 3:
        print("CAD连接失败，程序退出")
        exit()

    acad, space, doc = connect_result

    try:

        # 起点坐标，终点坐标
        contour_dict = {}
        current_contour_id = None
        current_points = []

        with open(locap, "r", encoding="utf-8") as f:
            #坐标转换
            for line in f:
                line_stripped = line.strip()
                # 识别标题行，提取轮廓编号
                if line_stripped.startswith("=== 帧") and line_stripped.endswith("==="):
                    # 保存上一个轮廓的坐标
                    if current_contour_id is not None and current_points and len(current_points) >= 3:
                        contour_dict[current_contour_id].append(current_points)
                    # 提取新的轮廓编号
                    current_contour_id = line_stripped.replace("=== 帧", "").replace("===", "").strip()
                    contour_dict[current_contour_id] = []
                    current_points = []
                    continue
                # 跳过空行
                elif not line_stripped:
                    if current_contour_id is not None and current_points and len(current_points) >= 3:
                        contour_dict[current_contour_id].append(current_points)
                    current_points = []
                    continue

                # 解析坐标行
                else:
                    try:
                        x, y = map(float, line_stripped.split(","))
                        # 过滤异常坐标（根据你的视频分辨率调整，比如1920x1080）
                        if 10 < x < 1920 and 10 < y < 1080:
                            current_points.append((x, y))
                    except:
                        print(f"跳过格式错误的行：{line_stripped}")
                        continue

            # 保存最后一个轮廓的坐标
            if current_contour_id is not None and current_points:
                contour_dict[current_contour_id] = current_points
         #绘制CAD图像
        sorted_frames = sorted(contour_dict.keys(), key=lambda x: int(x))
        #转换元组类型
        total_frames = len(sorted_frames)
        for id , contour_id in enumerate(sorted_frames,1):
            force_cad_active(doc)

            #points = contour_dict[contour_id]
            sub_contours = contour_dict[contour_id]
            frames = []
            try:
                # 遍历当前帧的每个子轮廓，单独绘制
                for points  in sub_contours:
                    # 过滤点数不足的轮廓
                    if len(points) < 3:
                        continue

                    point = []
                    for (x, y) in points:
                        point.extend([x, y, 0.0])
                    points_variant = win32com.client.VARIANT( pythoncom.VT_ARRAY | pythoncom.VT_R8, point )
                    # 绘制多段线

                    polyline = space.AddPolyline(points_variant)

                    # 设置线条属性
                    polyline.Color = 7
                    polyline.Lineweight = 50

                    first_point = points[0]
                    last_point = points[-1]
                    #print(first_point)
                    # 计算首尾点的距离
                    distance = ((first_point[0] - last_point[0]) ** 2 + (first_point[1] - last_point[1]) ** 2) ** 0.5
                    if distance < 1.5:
                        polyline.Closed = True
                    frames.append(polyline)
                    print("直线绘制成功")
                    #acad.ActiveDocument.SendCommand("_REGEN ")
                force_cad_active(doc)

                #强制缩放
                if id <2:
                    acad.ActiveDocument.Application.ZoomExtents()

            except Exception as e:
                print(f"绘图失败: {e}")
            #force_cad_active(doc)
            #time.sleep(0.03)

            try:
                for frame in frames:

                    frame.Delete()

                #acad.ActiveDocument.SendCommand("_REGEN ")
                #acad.ActiveDocument.Application.ZoomExtents()
            except Exception as del_e:
                continue
    finally:
        # 释放COM环境
        pythoncom.CoUninitialize()