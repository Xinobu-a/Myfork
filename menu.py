"""菜单系统模块"""
import os
class MenuSystem:
    def __init__(self,file_handler,formatter):
        self.fil_handler=file_handler
        self.formatter = formatter
        self.current_content=None
        self.current_file_path=None
        self.is_modified=False
    def show_menu(self):
        print("\n" + "=" * 35)
        print("     文档自动排版工具")
        print("=" * 35)
        print("1. 读取文档文件")
        print("2. 基础格式排版（默认格式 + 冗余清理）")
        print("3. 退出系统")
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
            stats = self.formatter.get_statistics(formatted_content)
            print(f"\n排版后统计:")
            print(f"  - 字符数: {stats['char_count']}")
            print(f"  - 单词数: {stats['word_count']}")
            print(f"  - 段落数: {stats['paragraph_count']}")
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
    def exit_system(self):
        if self.is_modified:
            print("\n文档尚未保存！")
            save_choice = input("是否保存当前文档？(y/n): ").strip().lower()
            if save_choice == 'y':
                self.save_document()
            print("\n感谢使用文档自动排版工具！再见！")
            return True
    def run(self):
        while True:
            self.show_menu()
            choice = input("\n请选择操作: ").strip()

            if choice == '1':
                self.read_document()
            elif choice == '2':
                self.basic_format()
            elif choice == '4':
                self.save_document()
            elif choice == '5':
                if self.exit_system():
                    break
            else:
                print("\n无效输入，请输入1-5之间的数字")


