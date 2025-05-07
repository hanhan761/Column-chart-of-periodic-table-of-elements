import pygame
import pandas as pd  # 新增pandas用于读取Excel
from pygame.locals import *
from OpenGL.GL import *
from OpenGL.GLU import *
from PIL import Image, ImageDraw, ImageFont


def setup_display():
    """初始化 Pygame 并设置 OpenGL 视角"""
    pygame.init()
    display = (1200, 800)
    pygame.display.gl_set_attribute(pygame.GL_MULTISAMPLEBUFFERS, 1)
    pygame.display.gl_set_attribute(pygame.GL_MULTISAMPLESAMPLES, 4)
    screen = pygame.display.set_mode(display, DOUBLEBUF | OPENGL)
    gluPerspective(45, (display[0] / display[1]), 0.1, 100.0)
    glTranslatef(0, 0, -30)
    glEnable(GL_DEPTH_TEST)
    glEnable(GL_LINE_SMOOTH)
    glHint(GL_LINE_SMOOTH_HINT, GL_NICEST)
    return display


def load_elements_from_excel(filename):
    """从Excel文件加载元素数据"""
    try:
        # 使用pandas读取Excel，默认读取第一个工作表
        df = pd.read_excel(filename, engine='openpyxl')  # 需要openpyxl引擎支持xlsx

        # 转换数据格式并构建元素列表
        elements = []
        for _, row in df.iterrows():
            elements.append({
                'symbol': str(row['元素']),  # 元素符号（字符串）
                'number': str(row['编号']),  # 元素编号（字符串）
                'x': float(row['x']),  # x坐标（浮点型）
                'y': float(row['y']),  # y坐标（浮点型）
                'width': float(row['长度']),  # 宽度（浮点型）
                'length': float(row['宽度']),  # 长度（浮点型）
                'height': float(row['高度'])  # 高度（浮点型）
            })
        return elements
    except FileNotFoundError:
        raise Exception(f"错误：未找到数据文件 {filename}")
    except Exception as e:
        raise Exception(f"读取Excel出错：{str(e)}")


def create_text_texture(element_number, element_symbol):
    """创建带有元素编号和符号的纹理（与原逻辑一致）"""
    width, height = 512, 512  # Increased texture resolution for sharper text
    image = Image.new('RGB', (width, height), color=(240, 240, 240))
    draw = ImageDraw.Draw(image)

    font_paths = ["arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                  "/System/Library/Fonts/SFNSDisplay.ttf"]
    used_font = None
    for path in font_paths:
        try:
            used_font = ImageFont.truetype(path, 100)  # Increased font size for better clarity
            break
        except (OSError, IOError):
            continue
    if not used_font:
        used_font = ImageFont.load_default()

    draw.text((30, 20), element_number, fill=(30, 30, 30), font=used_font)
    symbol_font = ImageFont.truetype(used_font.path, 160) if used_font.path else ImageFont.load_default()
    bbox = draw.textbbox((0, 0), element_symbol, font=symbol_font)
    symbol_width = bbox[2] - bbox[0]
    symbol_height = bbox[3] - bbox[1]
    x_pos = (width - symbol_width) // 2
    y_pos = (height - symbol_height) // 2 + 20
    draw.text((x_pos, y_pos), element_symbol, fill=(30, 30, 30), font=symbol_font)

    image = image.transpose(Image.FLIP_TOP_BOTTOM)
    image_data = image.convert("RGBA").tobytes()

    texture_id = glGenTextures(1)
    glBindTexture(GL_TEXTURE_2D, texture_id)
    glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR_MIPMAP_LINEAR)  # Improved texture filtering
    glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR)
    glTexParameteri(GL_TEXTURE_2D, GL_GENERATE_MIPMAP, GL_TRUE)
    glTexImage2D(GL_TEXTURE_2D, 0, GL_RGBA, width, height, 0, GL_RGBA, GL_UNSIGNED_BYTE, image_data)
    return texture_id


def draw_column(position, dimensions, element_number, element_symbol):
    """绘制柱子（与原逻辑一致）"""
    x, y, z = position
    width, length, height = dimensions

    vertices_bottom = [
        [x - width / 2, y - length / 2, z],
        [x + width / 2, y - length / 2, z],
        [x + width / 2, y + length / 2, z],
        [x - width / 2, y + length / 2, z]
    ]
    vertices_top = [
        [x - width / 2, y - length / 2, z + height],
        [x + width / 2, y - length / 2, z + height],
        [x + width / 2, y + length / 2, z + height],
        [x - width / 2, y + length / 2, z + height]
    ]

    glBegin(GL_QUADS)
    for i in range(4):
        next_i = (i + 1) % 4
        top_color = (0.3 + i * 0.05, 0.3 + i * 0.05, 0.9)
        bottom_color = (0.2, 0.2, 0.7)
        glColor3fv(bottom_color)
        glVertex3fv(vertices_bottom[i])
        glColor3fv(bottom_color)
        glVertex3fv(vertices_bottom[next_i])
        glColor3fv(top_color)
        glVertex3fv(vertices_top[next_i])
        glColor3fv(top_color)
        glVertex3fv(vertices_top[i])
    glEnd()

    texture_id = create_text_texture(element_number, element_symbol)
    glEnable(GL_TEXTURE_2D)
    glEnable(GL_BLEND)
    glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)
    glBindTexture(GL_TEXTURE_2D, texture_id)
    glBegin(GL_QUADS)
    glTexCoord2f(0, 0)
    glVertex3fv(vertices_top[0])
    glTexCoord2f(1, 0)
    glVertex3fv(vertices_top[1])
    glTexCoord2f(1, 1)
    glVertex3fv(vertices_top[2])
    glTexCoord2f(0, 1)
    glVertex3fv(vertices_top[3])
    glEnd()
    glDisable(GL_TEXTURE_2D)
    glDisable(GL_BLEND)

    glColor3fv((0, 0, 0))
    glLineWidth(2.5)  # Increased line width for better edge visibility
    for i in range(4):
        next_i = (i + 1) % 4
        glBegin(GL_LINE_STRIP)
        glVertex3fv(vertices_bottom[i])
        glVertex3fv(vertices_bottom[next_i])
        glVertex3fv(vertices_top[next_i])
        glVertex3fv(vertices_top[i])
        glVertex3fv(vertices_bottom[i])
        glEnd()


def main():
    """主函数：加载Excel数据并批量绘制"""
    display = setup_display()
    elements = load_elements_from_excel("data.xlsx")  # 改为读取xlsx文件

    x_rot = -33.6
    y_rot = -0.4
    mouse_down = False
    last_x, last_y = 0, 0
    zoom = -29
    pan_x, pan_y = 14.26, 7.44
    right_mouse_down = False

    while True:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                quit()
            elif event.type == pygame.MOUSEBUTTONDOWN:
                if event.button == 1:
                    mouse_down = True
                    last_x, last_y = event.pos
                elif event.button == 3:  # Right mouse button for panning
                    right_mouse_down = True
                    last_x, last_y = event.pos
                elif event.button == 4:  # Mouse wheel up
                    zoom = min(zoom + 1, -5)  # Clamp zoom to prevent excessive zoom-in
                elif event.button == 5:  # Mouse wheel down
                    zoom = max(zoom - 1, -50)  # Clamp zoom to prevent excessive zoom-out
            elif event.type == pygame.MOUSEBUTTONUP:
                if event.button == 1:
                    mouse_down = False
                elif event.button == 3:  # Right mouse button released
                    right_mouse_down = False
            elif event.type == pygame.MOUSEMOTION:
                if mouse_down:
                    dx = event.pos[0] - last_x
                    dy = event.pos[1] - last_y
                    x_rot += dy * 0.3
                    y_rot += dx * 0.3
                    last_x, last_y = event.pos
                elif right_mouse_down:
                    dx = event.pos[0] - last_x
                    dy = event.pos[1] - last_y
                    pan_x += dx * 0.01
                    pan_y -= dy * 0.01
                    last_x, last_y = event.pos
            elif event.type == pygame.MOUSEBUTTONDOWN:
                if event.button == 4:  # Mouse wheel up
                    zoom = min(zoom + 1, -5)  # Clamp zoom to prevent excessive zoom-in
                elif event.button == 5:  # Mouse wheel down
                    zoom = max(zoom - 1, -50)  # Clamp zoom to prevent excessive zoom-out

        glClearColor(0.95, 0.95, 0.95, 1.0)
        glClear(GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT)

        glLoadIdentity()
        gluPerspective(45, (display[0] / display[1]), 0.1, 200.0)
        glTranslatef(pan_x, pan_y, zoom)  # Apply pan offsets
        glRotatef(x_rot, 1, 0, 0)
        glRotatef(y_rot, 0, 1, 0)

        for elem in elements:
            position = (elem['x'] * 1.5, elem['y'] * 1.5, 0)
            dimensions = (elem['width'], elem['length'], elem['height'])
            draw_column(
                position=position,
                dimensions=dimensions,
                element_number=elem['number'],
                element_symbol=elem['symbol']
            )

        pygame.display.flip()
        pygame.time.wait(10)


if __name__ == "__main__":
    main()
