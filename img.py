import cv2
from pathlib import Path
path=Path(r"E:\Python\project\cv2\location\1.png")
read=cv2.imread(path)
gray=cv2.cvtColor(read,cv2.COLOR_BGR2GRAY)
cv2.imshow("original",gray)
cv2.waitKey(0)
cv2.destroyAllWindows()