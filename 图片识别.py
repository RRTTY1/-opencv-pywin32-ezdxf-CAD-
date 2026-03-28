import cv2
from pathlib import Path
import numpy as np
#存放路径
save=Path(r"E:\Python\project\cv2\frames")
#坐标路径
locap=Path(r"E:\Python\project\cv2\location")
save.mkdir(exist_ok=True,parents=True)
locap.mkdir(exist_ok=True,parents=True)
extensions=["*.jpg","*.png"]
img_path=[]
for ext in extensions:
    img_path.extend(save.glob(ext))
img_path=sorted(img_path)
if not img_path:
    print("未找到任何图片")
    exit()
print(f"共找到 {len(img_path)} 张图片，开始轮廓提取...\n")
output_txt = locap / "contour_coords.txt"
with open(output_txt, "w", encoding="utf-8") as f:
    for id , img in enumerate(img_path,1):
        img_loc = cv2.imdecode(np.fromfile(str(img), dtype=np.uint8), cv2.IMREAD_COLOR)
        if img_loc is None:
            continue
        gray=cv2.cvtColor(img_loc,cv2.COLOR_BGR2GRAY)
        binary = cv2.threshold(gray,127,255,cv2.THRESH_BINARY)[1]
        contours, hierarchy = cv2.findContours(binary, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        f.write(f"=== 帧 {id} ===\n")
        for cnt in contours:
            if cv2.contourArea(cnt) > 5:
                img_h = img_loc.shape[0]
                f.write(f"\n")
                for point in cnt:
                    x, y = point[0]
                    y_flipped = img_h - y
                    f.write(f"{x}, {y_flipped}\n")
        f.write("\n")
print("提取成功")