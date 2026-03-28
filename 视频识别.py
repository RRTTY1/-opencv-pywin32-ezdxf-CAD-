import cv2
from pathlib import Path
vc=cv2.VideoCapture("badapple.mp4")
#存放路径
save=Path(r"E:\Python\project\cv2\frames")
save.mkdir(exist_ok=True,parents=True)
if vc.isOpened():
    ret,frame=vc.read()
else:
    ret=False
frame_num=0
while ret:
    ret,frame=vc.read()
    if ret is None:
        break
    if ret==True:
        gray=cv2.cvtColor(frame,cv2.COLOR_BGR2GRAY)
        binary = cv2.threshold(gray,127,255,cv2.THRESH_BINARY,cv2.THRESH_BINARY_INV)[1]
        #canny = cv2.Canny(gray_blur, 100, 200)
        save_path = save / f"{frame_num:5d}.png"
        cv2.imencode('.png', binary)[1].tofile(str(save_path))
        frame_num += 1
        cv2.imshow("result",binary)
        if cv2.waitKey(30) & 0xFF == 27:
            break
vc.release()
cv2.destroyAllWindows()
print(f"共处理{frame_num}张图片")