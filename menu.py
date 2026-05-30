"""菜单系统模块"""
import os
class MenuSystem:
    def __init__(self,file_handler,formatter):
        self.file_handler=file_handler
        self.formatter = formatter
        self.current_content=None
        self.current_file_path=None
        self.is_modified=False
    def show_main_menu(self):
        print("\n" + "=" * 35)
        print("     文档自动排版工具")
        print("=" * 35)
        print("1. 读取文档文件")
        print("2. 基础格式排版（默认格式 + 冗余清理）")
        print("3. 进阶自定义操作")
        print("4. 保存排版后文档")
        print("5. 退出系统")
        print("-" * 35)
    def read_document(self):
        file_path = input("\n请输入文档文件路径: ").strip()
        try:
            content, file_type = self.file_handler.read_file(file_path)
            self.current_content = content
            self.current_file_path = file_path
            self.is_modified = False
            print(f"文档读取成功: {file_path}")
            print(f"文件类型: {file_type}")
            print(f"文档长度: {len(content)}")
        except Exception as e:
            print(f"错误: {str(e)}")
    def basic_format(self):
        if not self.current_content:
            print("\n错误: 请先读取文档文件")
            return
        print("\n正在进行基础排版...")
        print("  - 清理多余空格")
        print("  - 清理连续空行")
        print("  - 清理无意义换行")
        print("  - 应用默认格式（宋体小四、固定行距18磅）")
        try:
            formatted_content = self.formatter.apply_basic_format(self.current_content)
            self.current_content = formatted_content
            self.is_modified = True
            print("\n基础排版完成！")
        except Exception as e:
            print(f"\n排版失败: {str(e)}")
    def save_document(self):
        if not self.current_content:
            print("\n错误: 没有可保存的文档内容")
            return
        if self.current_file_path:
            base_name = os.path.splitext(self.current_file_path)[0]
            default_path = f"{base_name}_formatted.txt"
        else:
            default_path = "formatted_document.txt"
        file_path = input(f"\n请输入保存路径（直接回车使用 {default_path}）: ").strip()
        if not file_path:
            file_path = default_path
        print("\n选择保存格式:")
        print("1. 纯文本 (.txt)")
        print("2. Word文档 (.docx)")
        format_choice = input("请选择 (1/2，默认1): ").strip()
        try:
            if format_choice == '2':
                if not file_path.endswith('.docx'):
                    file_path = file_path.replace('.txt', '.docx')
                self.formatter.save_as_word(self.current_content, file_path)
                print(f"\n文档已保存为Word格式: {file_path}")
            else:
                if not file_path.endswith('.txt'):
                    file_path += '.txt'
                self.file_handler.save_file(self.current_content, file_path, 'text')
                print(f"\n文档已保存为文本格式: {file_path}")
            self.is_modified = False
        except Exception as e:
            print(f"\n保存失败: {str(e)}")
    def show_advanced_menu(self):
        """显示进阶菜单"""
        print("\n" + "=" * 35)
        print("     进阶自定义操作")
        print("=" * 35)
        print("1. 设置自定义格式参数")
        print("2. 查看当前格式设置")
        print("3. 文档字数统计")
        print("4. 返回主菜单")
        print("-" * 35)
    def advanced_operations(self):
        """进阶操作入口"""
        while True:
            self.show_advanced_menu()
            choice = input("\n请选择操作: ").strip()
            if choice == '1':
                self.set_custom_format()
            elif choice == '2':
                self.show_current_settings()
            elif choice == '3':
                self.show_statistics()
            elif choice == '4':
                break
            else:
                print("\n✗ 无效输入，请输入1-4之间的数字")
    def set_custom_format(self):
        """设置自定义格式"""
        if not self.current_content:
            print("\n✗ 错误: 请先读取文档文件")
            return
        print("\n当前格式设置:")
        current = self.formatter.get_current_settings()
        print(f"  字体: {current['font_name']}")
        print(f"  字号: {current['font_size']}pt")
        print(f"  行距: {current['line_spacing']}磅")
        print(f"  首行缩进: {current['first_line_indent']}英寸")
        print("\n请输入新的格式参数（直接回车保留当前值）:")
        font_name = input(f"字体 [{current['font_name']}]: ").strip()
        font_size = input(f"字号(磅) [{current['font_size']}]: ").strip()
        line_spacing = input(f"行距(磅) [{current['line_spacing']}]: ").strip()
        indent = input(f"首行缩进(英寸) [{current['first_line_indent']}]: ").strip()
        try:
            font_size = int(font_size) if font_size else None
            line_spacing = int(line_spacing) if line_spacing else None
            indent = float(indent) if indent else None
            # 直接修改自定义设置
            if font_name:
                self.formatter.custom_settings['font_name'] = font_name
            if font_size:
                self.formatter.custom_settings['font_size'] = font_size
            if line_spacing:
                self.formatter.custom_settings['line_spacing'] = line_spacing
            if indent:
                self.formatter.custom_settings['first_line_indent'] = indent

            # 重新应用格式
            if self.current_content:
                formatted_content = self.formatter.apply_basic_format(self.current_content)
                self.current_content = formatted_content
                self.is_modified = True
            print("\n✓ 自定义格式已更新！")
        except ValueError:
            print("\n✗ 错误: 请输入有效的数字")
    def show_current_settings(self):
        """显示当前格式设置"""
        settings = self.formatter.get_current_settings()
        print("\n当前格式设置:")
        print("=" * 30)
        print(f"  字体: {settings['font_name']}")
        print(f"  字号: {settings['font_size']}pt")
        print(f"  行距: {settings['line_spacing']}磅")
        print(f"  首行缩进: {settings['first_line_indent']}英寸")
        print("=" * 30)
    def show_statistics(self):
        """显示文档统计信息（字数统计）"""
        if not self.current_content:
            print("\n✗ 错误: 请先读取文档文件")
            return
        stats = self.formatter.get_statistics(self.current_content)
        print("\n文档统计信息:")
        print("=" * 30)
        print(f"  字符数（不含空格）: {stats['char_count']}")
        print(f"  单词数: {stats['word_count']}")
        print(f"  行数: {stats['line_count']}")
        print(f"  段落数: {stats['paragraph_count']}")
        print("=" * 30)
    def exit_system(self):
        if self.is_modified:
            print("\n文档尚未保存！")
            save_choice = input("是否保存当前文档？(y/n): ").strip().lower()
            if save_choice == 'y':
                self.save_document()
        print("\n感谢使用文档自动排版工具！再见！")
        return True
    def run(self):
        """运行菜单系统"""
        while True:
            self.show_main_menu()
            choice = input("\n请选择操作: ").strip()
            
            if choice == '1':
                self.read_document()
            elif choice == '2':
                self.basic_format()
            elif choice == '3':
                self.advanced_operations()
            elif choice == '4':
                self.save_document()
            elif choice == '5':
                if self.exit_system():
                    break
            else:
                print("\n✗ 无效输入，请输入1-5之间的数字")


