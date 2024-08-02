# import pythoncom
# import win32com.client
# #import math
#
# cad = win32com.client.Dispatch('autocad.application')
# doc = cad.activedocument
# modelspace = doc.modelspace

####Line
# start_point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(20,20,0))
# end_point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(100,100,0))
# modelspace.Addline(start_point,end_point)

####point
# doc.setvariable('PDMODE',3)
# doc.setvariable('PDSIZE',1)
# point1 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(10,10,0))
# modelspace.Addpoint(point1)

####polyline
# vertics = [1,1,0,10,10,0,20,1,0]
# v_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,vertics)
# aa = modelspace.AddPolyline(v_array)
# aa.Closed = True
#
# aa.color = 1
# aa.SetProperties(Lineweight=200)

####circle
# aa = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(20,20,0))
# modelspace.Addcircle(aa,5)

####Arc
# aa = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(10,20,0))
# bb = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(15,25,0))
# cc = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,(20,20,0))
# modelspace.AddArc(aa,bb,cc)

####Rectangle
# jihe = [0,0,0,10,0,0,10,5,0,0,5,0,0,0,0]
# points = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, jihe)
# modelspace.AddPolyline(points)

####Polygon
# import win32com.client
# import pythoncom
# import math
#
# # 启动AutoCAD应用程序
# acad = win32com.client.Dispatch("AutoCAD.Application")
# acad.Visible = True
#
# # 获取Active Document
# doc = acad.ActiveDocument
# model_space = doc.ModelSpace
#
# def multi_poly(center,radius,num_sides):
# # 创建一个多边形的中心点和半径
# # center = (10, 10, 0)
# # radius = 5
# # num_sides = 6
#
#     # 计算多边形的顶点坐标
#     points = []
#     for i in range(num_sides):
#         angle = 2 * math.pi * i / num_sides
#         x = center[0] + radius * math.cos(angle)
#         y = center[1] + radius * math.sin(angle)
#         points.extend([x, y, 0])  # 每个顶点包含x, y, z坐标
#
#     # 为了闭合多边形，添加第一个顶点到最后
#     points.extend(points[:3])
#
#     # 将顶点坐标转换为VARIANT数组格式
#     points_variant = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, points)
#
#     # 创建多边形对象
#     polygon = model_space.AddPolyline(points_variant)
#
# multi_poly((10, 10, 0),5,6)

####Ellipse
# center_points = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (0,0,0))
# major_axis = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (10,0,0))
# ratio = 0.5
# modelspace.AddEllipse(center_points,major_axis,ratio)

####Elliptical arc
# center_points = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (0,0,0))
# major_axis = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (10,0,0))
# ratio = 0.5
#
# ellipse = modelspace.AddEllipse(center_points,major_axis,ratio)
#
# ellipse.StartAngle = 0
# ellipse.EndAngle = 1.57

####Figure creating and filling
# vertices = [0,0,0,10,0,0,10,5,0,0,5,0,0,0,0]
# points = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, vertices)
# rect = modelspace.AddPolyline(points)
#
# hatch = modelspace.AddHatch(0,'SOLID',True)
#
# hatch.AppendOuterLoop(win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [rect]))

####Figure moving and copy
# center_point1 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (50,50,0))
# rudis = 10
# circle1 = modelspace.AddCircle(center_point1,rudis)
# center_point2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (25,25,0))
# circle1.copy()
# circle1.move(center_point1,center_point2)

####Figure delete,rotate,mirror
# point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (50,50,0))
# text = modelspace.Addtext('abcd',point,5)
# point1 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (0,0,0))
# point2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (43,43,0))
# text_mirror = text.mirror(point1,point2)
# rotate_point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (80,80,0))
# text_mirror.rotate(rotate_point,3)
# text_mirror.delete()

####Figure scale
# point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (50,50,0))
# text = modelspace.Addtext('abcd',point,5)
# jizhun_point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (24,50,0))
# text.Scaleentity(jizhun_point,2)

####Array
# point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (50,50,0))
# radius =10
# circle = modelspace.AddCircle(point,10)
# circle.ArrayRectangular(5,5,1,15,15,0)

####single line
# point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (80,80,0))
# text = modelspace.AddText('aaaa',point,12)
# text.Alignment = 1
# text.TextAlignmentPoint = point1
# text.Rotation = 1.57
# modelspace.AddCircle(point,10)

####multiple line
# point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (80,80,0))
# textString = 'hello. nice to meet you'
# text = modelspace.AddMText(point,5,textString)
# text.Height = 10

####liner and align
# point1 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (0,0,0))
# point2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (10,0,0))
# point3 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (5,2,0))
#
# aaa = modelspace.AddDimAligned(point1,point2,point3)
#
# aaa.TextHeight = 1
# aaa.ArrowheadSize =1

# point1 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (0,0,0))
# point2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (10,5,0))
# point3 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (5,7,0))
#
# aaa = modelspace.AddDimAligned(point1,point2,point3)

####leader
# points = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (10, 10, 0, 35, 35, 0))
# point2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (35, 45, 0))
#
# text = modelspace.AddMText(point2, 5, 'asfd')
#
# leader = modelspace.AddLeader(points, text, 2)
#
# #leader.ArrowHeadType =1
# leader.ArrowHeadSize = 3
# leader.Update()

####insert and design table
# point = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (200, 200, 0))
# table = modelspace.AddTable(point,5,5,20,20)
# point1 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (10, 10, 0))
# point2 = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (50, 50, 0))
# table_copy = table.copy()
# table_copy.move(point1,point2)
#table.InsertRows(1,10,3)
#table.mergecells(2,3,1,2)
# table.settext(0,0,'nice to meet you')
# table.settext(2,2,'how are you')
# table.setalignment(4,7)
# table.delete()
# table_copy.update()

####layer
# layers = doc.Layers
# new_layer = layers.Add('test')
# new_layer.color =1
# new_layer.lineweight =35
# new_layer.linetype = 'Continuous'
# doc.activelayer = new_layer

####Blocks
git --version