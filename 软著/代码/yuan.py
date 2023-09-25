import matplotlib.pyplot as plt
from matplotlib.patches import Circle
import random

def generate_disjoint_circles(rect_width, rect_height, num_circles, desired_area_ratio):
    circles = []
    total_area = rect_width * rect_height
    target_area = total_area * desired_area_ratio
    current_area = 0.0

    while len(circles) < num_circles:
        radius = random.uniform(0, min(rect_width, rect_height) / 2)
        x = random.uniform(radius, rect_width - radius)
        y = random.uniform(radius, rect_height - radius)
        circle = Circle((x, y), radius)

        # 检查圆形与现有圆形是否相交
        intersect = False
        for existing_circle in circles:
            if circle.intersects(existing_circle):
                intersect = True
                break

        # 如果相交或超过目标面积，则重新生成圆形
        if intersect or current_area + circle.area > target_area:
            continue

        # 添加圆形到列表
        circles.append(circle)
        current_area += circle.area

    return circles

# 定义矩形的宽度和高度
rect_width = 10
rect_height = 8

# 定义生成圆形的个数和目标面积占比
num_circles = 5
desired_area_ratio = 0.4

# 生成不相交的圆形
circles = generate_disjoint_circles(rect_width, rect_height, num_circles, desired_area_ratio)

# 创建图形和坐标轴
fig, ax = plt.subplots()

# 绘制矩形
rect = plt.Rectangle((0, 0), rect_width, rect_height, fill=False)
ax.add_patch(rect)

# 绘制圆形
for circle in circles:
    ax.add_patch(circle)

# 设置坐标轴范围
ax.set_xlim(0, rect_width)
ax.set_ylim(0, rect_height)

# 显示图形
plt.axis('equal')  # 使x和y轴的单位比例相等
plt.show()
