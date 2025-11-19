import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.patches import FancyBboxPatch
import numpy as np
import pandas as pd
from matplotlib.gridspec import GridSpec
import matplotlib.colors as mcolors


# 从data.txt中提取数据
def parse_data_file():
    data = {
        'features': [],
        'nominal': [],
        'actual': []
    }

    # 模拟从文件中读取的数据
    feature_data = [
        ('DIST-1', 499.883160000000, 499.884468244403),
        ('DIST-2', 499.883160000000, 499.884005549235),
        ('DIST-3', 499.883160000000, 499.884108367401),
        ('DIST-4', 379.916150000000, 379.916343668494),
        ('DIST-5', 379.916150000000, 379.916375631849),
        ('DIST-6', 379.916150000000, 379.916506951887),
        ('DIST-7', 259.990730000000, 259.991133422226),
        ('DIST-8', 259.990730000000, 259.991280335469),
        ('DIST-9', 259.990730000000, 259.991266691397),
        ('DIST-10', 139.963300000000, 139.963970368656),
        ('DIST-11', 139.963300000000, 139.964079013519),
        ('DIST-12', 139.963300000000, 139.964163007178),
        ('DIST-13', 20.000450000000, 20.000835129703),
        ('DIST-14', 20.000450000000, 20.000859332084),
        ('DIST-15', 20.000450000000, 20.000864971181)
    ]

    for feature, nominal, actual in feature_data:
        data['features'].append(feature)
        data['nominal'].append(nominal)
        data['actual'].append(actual)

    return data


# 计算偏差
def calculate_deviations(data):
    deviations = []
    for i in range(len(data['nominal'])):
        deviation = (data['actual'][i] - data['nominal'][i]) * 1000  # 转换为微米
        deviations.append(deviation)
    return deviations


# 创建图表
def create_visualization(data):
    deviations = calculate_deviations(data)

    # 设置颜色
    colors = {
        'background': '#F5F5F5',
        'header': '#2E4057',
        'text': '#333333',
        'grid': '#CCCCCC',
        'positive': '#4A90E2',
        'negative': '#E74C3C',
        'table_header': '#34495E',
        'table_row1': '#ECF0F1',
        'table_row2': '#FFFFFF'
    }

    # 创建图形
    fig = plt.figure(figsize=(12, 10), facecolor=colors['background'])
    gs = GridSpec(2, 1, height_ratios=[2, 1], hspace=0.3)

    # 上部分：偏差图
    ax1 = fig.add_subplot(gs[0])

    # 绘制偏差点
    x_positions = np.arange(1, len(deviations) + 1)

    # 根据偏差正负使用不同颜色
    positive_mask = np.array(deviations) >= 0
    negative_mask = np.array(deviations) < 0

    ax1.scatter(x_positions[positive_mask],
                np.array(deviations)[positive_mask],
                color=colors['positive'], s=80, zorder=5, label='Positive Deviation')
    ax1.scatter(x_positions[negative_mask],
                np.array(deviations)[negative_mask],
                color=colors['negative'], s=80, zorder=5, label='Negative Deviation')

    # 添加连接线
    ax1.plot(x_positions, deviations, color='gray', alpha=0.5, linewidth=1, zorder=4)

    # 设置Y轴范围
    max_dev = max(abs(min(deviations)), abs(max(deviations)))
    y_margin = max_dev * 0.2
    ax1.set_ylim(-max_dev - y_margin, max_dev + y_margin)

    # 添加零线
    ax1.axhline(y=0, color='black', linestyle='-', alpha=0.3, linewidth=1)

    # 添加网格
    ax1.grid(True, alpha=0.3, linestyle='-', linewidth=0.5)

    # 设置标签和标题
    ax1.set_xlabel('Measurement Point', fontsize=12, color=colors['text'], fontweight='bold')
    ax1.set_ylabel('Deviation (µm)', fontsize=12, color=colors['text'], fontweight='bold')
    ax1.set_title('Distance Measurement Deviations', fontsize=14, color=colors['header'], fontweight='bold', pad=20)

    # 设置X轴刻度
    ax1.set_xticks(x_positions)
    ax1.set_xticklabels([f'DIST-{i + 1}' for i in range(len(deviations))], rotation=45)

    # 添加图例
    ax1.legend(loc='upper right', framealpha=0.9)

    # 在右侧添加导航刻度
    ax1_right = ax1.twinx()
    nav_ticks = [-4.0, -3.2, -2.4, -1.6, -0.8, 0]
    ax1_right.set_ylim(ax1.get_ylim())
    ax1_right.set_yticks(nav_ticks)
    ax1_right.set_ylabel('Navigation', fontsize=10, color=colors['text'])
    ax1_right.tick_params(axis='y', labelsize=8)

    # 添加测量信息文本
    info_text = f"No. of distances: {len(deviations)}\nMeasuring Direction: X Direction\nCMM: SPECTRUM2 (171503)"
    ax1.text(0.02, 0.98, info_text, transform=ax1.transAxes, fontsize=10,
             verticalalignment='top', bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))

    # 下部分：表格
    ax2 = fig.add_subplot(gs[1])
    ax2.axis('off')

    # 准备表格数据
    nominal_values = [20.000450, 139.963300, 259.990730, 379.916150, 499.883160]
    table_data = []

    for nominal in nominal_values:
        # 找到对应的测量点
        indices = [i for i, n in enumerate(data['nominal']) if abs(n - nominal) < 0.001]
        actual_values = [data['actual'][i] for i in indices]
        deviations_group = [deviations[i] for i in indices]

        mean_actual = np.mean(actual_values)
        mean_deviation = np.mean(deviations_group)
        min_deviation = min(deviations_group)
        max_deviation = max(deviations_group)

        table_data.append([
            f"{nominal:.4f}",
            f"{mean_actual:.4f}",
            f"{mean_deviation:.4f}",
            f"{min_deviation:.4f}",
            f"{max_deviation:.4f}"
        ])

    # 创建表格
    columns = ['Nominal Value', 'Actual Value', 'Mean Deviation', 'Minimum', 'Maximum']

    table = ax2.table(cellText=table_data,
                      colLabels=columns,
                      cellLoc='center',
                      loc='center',
                      bbox=[0.1, 0.1, 0.8, 0.8])

    # 格式化表格
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.8)

    # 设置表格样式
    for i, key in enumerate(table.get_celld().keys()):
        cell = table.get_celld()[key]
        if key[0] == 0:  # 表头
            cell.set_facecolor(colors['table_header'])
            cell.set_text_props(color='white', weight='bold')
        else:
            if key[0] % 2 == 1:
                cell.set_facecolor(colors['table_row1'])
            else:
                cell.set_facecolor(colors['table_row2'])

    # 添加表格标题
    ax2.text(0.5, 0.95, 'Measurement Statistics', transform=ax2.transAxes,
             fontsize=12, color=colors['header'], fontweight='bold',
             horizontalalignment='center')

    # 添加底部信息
    bottom_info = "Temperature: 20.00/20.00/20.00/20.30 °C | Position: 669.37/-513.75/-55.59 mm | MPE = 2.1 + L/250 µm"
    fig.text(0.5, 0.02, bottom_info, ha='center', fontsize=9, color=colors['text'])

    plt.tight_layout()
    return fig


# 主程序
if __name__ == "__main__":
    # 解析数据
    data = parse_data_file()

    # 创建可视化
    fig = create_visualization(data)

    # 显示图表
    plt.show()

    # 可选：保存图表
    # fig.savefig('measurement_analysis.png', dpi=300, bbox_inches='tight')
    # print("图表已保存为 'measurement_analysis.png'")