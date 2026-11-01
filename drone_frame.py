import math
import os
import win32com.client

SAVE_DIR = r"C:\sw_project\parts"


def find_plane(model):
    """Возвращает первую плоскость в дереве построения (Front Plane)."""
    feat = model.FirstFeature
    while feat is not None:
        if feat.GetTypeName2 == "RefPlane":
            return feat
        feat = feat.GetNextFeature
    return None


def extrude(model, depth):
    """Выдавливает текущий эскиз на глубину depth (в метрах)."""
    model.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, depth, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, 0, 0, False)


# НОВОЕ: в конце списка параметров появился hole_d
def make_frame(name, n_arms, arm_length, arm_width, center_d, thickness, start_angle, hole_d):
    """Строит раму: центральная пластина, лучи и отверстия под моторы. Размеры в метрах, угол в градусах."""
    sw = win32com.client.Dispatch("SldWorks.Application")
    sw.Visible = True

    template = sw.GetUserPreferenceStringValue(8)
    model = sw.NewDocument(template, 0, 0, 0)

    # 1. Центральная пластина: круг
    find_plane(model).Select2(False, 0)
    model.SketchManager.InsertSketch(True)
    model.SketchManager.CreateCircleByRadius(0, 0, 0, center_d / 2)
    extrude(model, thickness)

    # 2. Лучи: каждый луч это повёрнутый прямоугольник
    for i in range(n_arms):
        angle = math.radians(start_angle + i * 360 / n_arms)
        dx, dy = math.cos(angle), math.sin(angle)   # куда направлен луч
        nx, ny = -dy, dx                            # перпендикуляр к лучу
        h = arm_width / 2

        # четыре угла прямоугольника
        x1, y1 = nx * h, ny * h
        x2, y2 = dx * arm_length + nx * h, dy * arm_length + ny * h
        x3, y3 = dx * arm_length - nx * h, dy * arm_length - ny * h
        x4, y4 = -nx * h, -ny * h

        find_plane(model).Select2(False, 0)
        model.SketchManager.InsertSketch(True)
        model.SketchManager.CreateLine(x1, y1, 0, x2, y2, 0)
        model.SketchManager.CreateLine(x2, y2, 0, x3, y3, 0)
        model.SketchManager.CreateLine(x3, y3, 0, x4, y4, 0)
        model.SketchManager.CreateLine(x4, y4, 0, x1, y1, 0)

        # НОВОЕ: отверстие под мотор в центре конца луча
        cx = dx * (arm_length - arm_width / 2)
        cy = dy * (arm_length - arm_width / 2)
        model.SketchManager.CreateCircleByRadius(cx, cy, 0, hole_d / 2)

        extrude(model, thickness)

    model.ViewZoomtofit2

    # Сохранение
    os.makedirs(SAVE_DIR, exist_ok=True)
    path = os.path.join(SAVE_DIR, name + ".SLDPRT")
    error = model.SaveAs3(path, 0, 1)
    if error == 0:
        print("Сохранено:", path)
    else:
        print("Не удалось сохранить, код ошибки:", error)


# НОВОЕ: запуск с диаметром отверстия 8 мм в конце списка
# имя, число лучей, длина луча, ширина луча, диаметр центра, толщина, начальный угол, диаметр отверстия
make_frame("frame_3arms", 3, 0.12, 0.02, 0.06, 0.003, 90, 0.008)
make_frame("frame_6arms", 6, 0.12, 0.02, 0.06, 0.003, 0, 0.008)
make_frame("frame_8arms", 8, 0.12, 0.02, 0.06, 0.003, 0, 0.008)
print("Готово")