# coding: utf8
from Autodesk.Revit.UI import IExternalEventHandler,ExternalEvent,TaskDialog,Selection
import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI

# from math import acos,pi
# from math import sqrt
# from decimal import Decimal,getcontext
 
# getcontext().prec = 30
 
# class Vector(object):
#     CANNOT_NORMALIZE_ZERO_VECTOR_MSG = 'Cannot normalize the zero vector'
#     def __init__(self, coordinates):
#         try:
#             if not coordinates:
#                 raise ValueError
#             self.coordinates = tuple([Decimal(x) for x in coordinates])
#             self.dimension = len(coordinates)
#         except ValueError:
#             raise ValueError('The coordinates must be nonempty')
#         except TypeError:
#             raise TypeError('The coordinates must be an iterable')
 
 
#     def __str__(self):
#         return 'Vector: {}'.format(self.coordinates)
#     #两个向量是否相等
 
#     def __eq__(self, v):
#         return self.coordinates == v.coordinates
 
#     # 加法
#     def plus(self,v):
#         return Vector([x + y for x,y in zip(self.coordinates,v.coordinates)])
 
#     # 减法
#     def minus(self,v):
#         return Vector([x - y for x,y in zip(self.coordinates,v.coordinates)])
 
#     # 向量的倍数
#     def times_scalar(self,m):
#         return Vector([Decimal(m)*x for x in self.coordinates])
 

#     def magnitude(self):
#         coordinates_squared = [x ** 2 for x in self.coordinates]
#         return sqrt(sum(coordinates_squared))
#     # 单位向量
#     def normalized(self):
#         try:
#             magnitude = self.magnitude()
#             return self.times_scalar(1.0/magnitude)
 
 
#         except ZeroDivisionError:
#             raise Exception(self.CANNOT_NORMALIZE_ZERO_VECTOR_MSG)   
 
#     # 两个向量的点积
#     def dot(self,v):
#         return sum([x*y for x,y in zip(self.coordinates,v.coordinates)])
 
#     # 两个向量之间的角度
#     def angle_with(self, v, in_degrees=False):
#         try:
#             u1 = self.normalized()
#             u2 = v.normalized()
#             dots = u1.dot(u2)
#             if abs(abs(dots) - 1) < 1e-10:
#                 if dots < 0:
#                     dots = -1
#                 else:
#                     dots = 1
#             angle_in_radians = acos(dots)
 
#             if in_degrees:
#                 degrees_per_radian = 180. / pi
#                 return angle_in_radians * degrees_per_radian
#             else:
#                 return angle_in_radians
 
 
#         except Exception as e:
#             if str(e) == self.CANNOT_NORMALIZE_ZERO_VECTOR_MSG:
#                 raise Exception('Cannot compute an angle with the zero vector')
#             else:
#                 raise e 
 
#     # 判断两个向量是否正交
#     def is_orthogonal_to(self, v, tolerance=1e-10):
#         return abs(self.dot(v)) < tolerance
 
 
#     # 是否是零向量
 
#     def is_zero(self, tolerance=1e-10):
#         return self.magnitude() < tolerance
 
#     # 两个向量是否平行
#     def is_parallel_to(self,v):
#         return (self.is_zero() or
#                 v.is_zero() or
#                 self.angle_with(v) == 0 or
#                 self.angle_with(v) == pi)
 
#     # 向量在另一个向量上的投影
#     def component_parallel_to(self,basis):
#         try:
#             u = basis.normalized()
#             weight = self.dot(u)
#             return u.times_scalar(weight)
#         except Exception as e:
#             if str(e) == self.CANNOT_NORMALIZE_ZERO_VECTOR_MSG:
#                 raise Exception('Cannot compute an angle with the zero vector')
#             else:
#                 raise e  
 
#     # 向量相对于投影向量的垂直向量
#     def component_orthogonal_to(self, basis):
#         try:
#             projection = self.component_parallel_to(basis)
#             return self.minus(projection)
#         except Exception as e:
#             if str(e) == self.CANNOT_NORMALIZE_ZERO_VECTOR_MSG:
#                 raise Exception('Cannot compute an angle with the zero vector')
#             else:
#                 raise e  
 
#     # 计算三维向量的向量积
#     def cross(self,v):
#         try:
#             x_1, y_1, z_1 = self.coordinates
#             x_2, y_2, z_2 = v.coordinates
#             new_coordinates = [ y_1 * z_2 - y_2 * z_1,
#                                 -(x_1 * z_2 - x_2 * z_1),
#                                 x_1 * y_2 - x_2 * y_1]
#             return Vector(new_coordinates)
#         except ValueError as e:
#             msg = str(e)
#             if msg == 'need more than 2 values to unpack':
#                 self_embedded_in_R3 = Vector(self.coordinates + ('0',))
#                 v_embedded_in_R3 = Vector(v.coordinates + ('0',))
#                 return self_embedded_in_R3.cross(v_embedded_in_R3)
#             elif (msg == 'too many values to unpack' or
#                 msg == 'need more than 1 value to unpack' ):
#                 raise Exception('wrong value number')
#         else:
#             raise e
 
#     # 两个向量组成的平行四边形 面积
#     def area_of_parallelogram_with(self,v):
#         cross_product = self.cross(v)
#         return cross_product.magnitude()
 
#     # 两个向量组成的三角形 面积
#     def area_of_triangle_with(self,v):
#         cross_product = self.cross(v)
#         return cross_product.magnitude()/2
 
# #例子
 
# v = Vector(['8.462','7.893','-8.187'])
# w = Vector(['6.984','-5.975','4.778'])
 
# print v.cross(w)


class SHANG(IExternalEventHandler):
    def __init__(self):
        self.zahl = 0
        
    def Execute(self,app):
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document

        view = uidoc.ActiveView

        direction = view.ViewDirection

        re1 = uidoc.Selection.PickObject(UI.Selection.ObjectType.Edge)

        re2 = uidoc.Selection.PickObject(UI.Selection.ObjectType.Edge)

        # re3 = uidoc.Selection.PickObject(UI.Selection.ObjectType.Edge)

        # re4 = uidoc.Selection.PickObject(UI.Selection.ObjectType.Edge)

        line_re1 = self.Get_Line_On_View(re1,doc,direction)

        Line_re2 = self.Get_Line_On_View(re2,doc,direction)

        Point_re2_re1 = self.Get_Point_On_Line(Line_re2.GetEndPoint(0),line_re1)

        print(Point_re2_re1)

        print(Point_re2_re1,Line_re2.GetEndPoint(0))

        Vek_re1_re2 = self.Get_Vektor(Point_re2_re1,Line_re2.GetEndPoint(0))

        print(Point_re2_re1,Line_re2.GetEndPoint(0))

        print(self.Get_Length(Vek_re1_re2))
        # e_r1 = doc.GetElement(re1.ElementId).GetGeometryObjectFromReference(re1)
        # curve_r1 = e_r1.AsCurve()
        # P0_r1 = curve_r1.GetEndPoint(0)
        # P1_r1 = curve_r1.GetEndPoint(1)

        
        

        # line1 = DB.Line.CreateBound(curve_r1.GetEndPoint(0), curve_r1.GetEndPoint(1))
        # print(curve_r1.GetEndPoint(0),curve_r1.GetEndPoint(1))


        
        # cl = uidoc.Selection.GetElementIds()
        # xyz = DB.XYZ(0,0,self.zahl/304.8)

        # t = DB.Transaction(doc,'move')
        # t.Start()

        # DB.ElementTransformUtils.MoveElements(doc,cl,xyz)
        # t.Commit()

    def GetName(self):
        return "SHANG"

    def Get_Point_OnView(self,Point_0,Direction):
        A = Direction.X
        B = Direction.Y
        C = Direction.Z

        X = Point_0.X
        Y = Point_0.Y
        Z = Point_0.Z

        t = (A*X+B*Y+Z*C)/(A*A+B*B+C*C)

        X0 = X - A*t
        Y0 = Y - B*t
        Z0 = Z - C*t

        return DB.XYZ(X0,Y0,Z0)

    def Get_Line_On_View(self,re,doc,direction):
        curve_r = doc.GetElement(re.ElementId).GetGeometryObjectFromReference(re).AsCurve()
        p0 = self.Get_Point_OnView(curve_r.GetEndPoint(0),direction)
        p1 = self.Get_Point_OnView(curve_r.GetEndPoint(1),direction)
        # print(curve_r.GetEndPoint(0),p0)
        # print(curve_r.GetEndPoint(1),p1)
        return DB.Line.CreateBound(p0,p1)
    
    def Get_Point_On_Line(self,Point_0,Line):
        a0 = Point_0.X
        b0 = Point_0.Y
        c0 = Point_0.Z
        a1 = Line.GetEndPoint(0).X
        b1 = Line.GetEndPoint(0).Y
        c1 = Line.GetEndPoint(0).Z
        a2 = Line.GetEndPoint(1).X
        b2 = Line.GetEndPoint(1).Y
        c2 = Line.GetEndPoint(1).Z

        A_Face = (b1-b0)*(c2-c0) - (c1-c0)*(b2-b0) 
        B_Face = (c1-c0)*(a2-a0) - (a1-a0)*(c2-c0) 
        C_Face = (a1-a0)*(b2-b0) - (b1-b0)*(a2-a0) 
        D_Face = a0*(c1-c0)*(b2-b0) + b0*(a1-a0)*(c2-c0) + c0*(b1-b0)*(a2-a0) - a0*(b1-b0)*(c2-c0) - b0*(c1-c0)*(a2-a0) - c0*(a1-a0)*(b2-b0)

        A_Line = a2-a1
        B_Line = b2-b1
        C_Line = c2-c1

        print(a0,b0,c0,a1,b1,c1,a2,b2,c2)
        print(A_Face,B_Face,C_Face,D_Face)
        print(A_Line,B_Line,C_Line)

        print(A_Face*A_Line+B_Face*B_Line+C_Face*C_Line)

        t = 0 - (A_Face*a1+B_Face*b1+C_Face*c1+D_Face)/(A_Face*A_Line+B_Face*B_Line+C_Face*C_Line)

        P_Out_X = A_Line * t  + a1
        P_Out_Y = B_Line * t  + b1
        P_Out_Z = C_Line * t  + c1

        return DB.XYZ(P_Out_X,P_Out_Y,P_Out_Z)

    def Get_Vektor(self,Point_0,Point_1):
        return DB.Line.CreateBound(Point_0,Point_1)
    
    def Get_Length(self,Line):
        return Line.Length * 304.8