import ezdxf
from pathlib import Path

# 读取坐标文件
locap = Path(r"C:\Users\LUKAS\Desktop\img\location\contour_coords.txt")
contour_dict = {}
current_contour_id = None
current_points = []

with open(locap, "r", encoding="utf-8") as f:
    for line in f:
        line_stripped = line.strip()
        if line_stripped.startswith("=== 帧") and line_stripped.endswith("==="):
            if current_contour_id is not None and current_points:
                contour_dict[current_contour_id] = current_points
            current_contour_id = line_stripped.replace("=== 帧", "").replace("===", "").strip()
            current_points = []
            continue
        elif not line_stripped:
            continue
        else:
            try:
                x, y = map(float, line_stripped.split(","))
                current_points.append((x, y))
            except:
                continue
    if current_contour_id is not None and current_points:
        contour_dict[current_contour_id] = current_points
doc = ezdxf.new(dxfversion='R2018')
msp = doc.modelspace()

for contour_id, points in contour_dict.items():
    if len(points) < 2:
        continue
    msp.add_lwpolyline(points, close=True, dxfattribs={'color': 7, 'lineweight': 30})

save_path = r"C:\Users\LUKAS\Desktop\img\location\轮廓结果.dxf"
doc.saveas(save_path)
print(f"DXF文件生成成功，路径：{save_path}")
