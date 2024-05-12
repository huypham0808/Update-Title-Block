from pyautocad import Autocad, APoint

acad = Autocad()
#p1 = APoint(10)
#p2 = APoint(20)
#line = a.model.AddLine(p1, p2)
#line = a.model.AddLine(p1, APoint(100, 200))
x1 = int(input("Nhap diem x1? "))
y1 = int(input("Nhap diem y1? "))
x2 = int(input("Nhap diem x2? "))
y2 = int(input("Nhap diem y2? "))
line1 = acad.model.AddLine(APoint(x1, y1), APoint(x2, y2))
maMau = int(input("Nhap ma mau: "))
line1.color = maMau
acad.app.ZoomExtents()
acad.prompt("Ve line 1 xong")
acad.prompt("Ban dang su dung layer " + line1.layer)