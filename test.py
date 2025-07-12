import tkinter as tk
from tkinter import colorchooser
from tkinter import ttk
from tkintertable import TableCanvas, TableModel

def create_table_window():
    # 创建顶级窗口
    window = tk.Toplevel()
    window.title("颜色配置表")
    window.geometry("600x400")

    # 创建表格框架
    table_frame = tk.Frame(window)
    table_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # 创建表格模型 - 使用更可靠的方法
    model = TableModel()

    # 设置列名 - 直接赋值给columnNames属性
    columns = ['Curve Color', 'Marker Color', 'Text Color']
    model.columnNames = columns  # 直接赋值为列表

    # 添加初始数据行（全部为黑色）- 使用createRow方法
    row_data = {col: "#000000" for col in columns}
    model.createRow(0, **row_data)  # 使用createRow代替addRow

    # 创建表格画布
    table = TableCanvas(
        table_frame,
        model=model,
        rowheight=40,
        cellwidth=150,
        read_only=True  # 禁止直接编辑
    )
    table.show()

    # 设置初始单元格背景色
    for col_idx in range(len(columns)):
        table.setCellBackground(0, col_idx, "#000000")

    # 双击事件处理函数
    def handle_double_click(event):
        # 获取点击的行列
        row = table.get_row_clicked(event)
        col = table.get_col_clicked(event)

        # 只处理前三列
        if row is not None and col < 3:
            # 获取当前颜色值
            col_name = columns[col]
            current_color = table.model.getValueAt(row, col_name)

            # 打开颜色选择器
            color_code = colorchooser.askcolor(
                initialcolor=current_color,
                title=f"选择 {col_name} 颜色"
            )

            # 如果用户选择了颜色
            if color_code and color_code[1]:
                new_color = color_code[1]
                # 更新单元格值和背景色
                table.model.setValueAt(row, col_name, new_color)
                table.setCellBackground(row, col, new_color)
                table.redraw()

    # 绑定双击事件
    table.bind("<Double-1>", handle_double_click)

    # 添加测试按钮
    test_frame = tk.Frame(window)
    test_frame.pack(pady=10)

    test_btn = ttk.Button(
        test_frame,
        text="测试颜色选择",
        command=lambda: handle_double_click(None)
    )
    test_btn.pack()

    return window

# 测试代码
if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("300x100")
    root.title("表格测试")

    # 安装验证
    try:
        from tkintertable import TableCanvas, TableModel
        # 测试表格功能
        test_model = TableModel()
        test_model.columnNames = ["Test1", "Test2"]
        test_model.createRow(0, Test1="val1", Test2="val2")
    except Exception as e:
        install_frame = tk.Frame(root)
        install_frame.pack(pady=20)

        error_label = tk.Label(install_frame, text=f"库初始化错误: {str(e)}", fg="red")
        error_label.pack(pady=5)

        tk.Label(install_frame, text="请尝试安装或更新 tkintertable").pack(pady=5)

        install_btn = ttk.Button(
            install_frame, 
            text="安装 tkintertable",
            command=lambda: __import__('os').system('pip install tkintertable')
        )
        install_btn.pack(pady=5)

        update_btn = ttk.Button(
            install_frame, 
            text="更新 tkintertable",
            command=lambda: __import__('os').system('pip install --upgrade tkintertable')
        )
        update_btn.pack(pady=5)
    else:
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=30)

        btn = ttk.Button(
            btn_frame, 
            text="打开颜色表格",
            command=create_table_window
        )
        btn.pack()

    root.mainloop()