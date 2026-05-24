"""主程序"""
from file_handler import FileHandler
from formatter import DocumentFormatter
from menu import MenuSystem
def main():
    try:
        file_handler = FileHandler()
        formatter = DocumentFormatter()
        menu = MenuSystem(file_handler, formatter)
        menu.run()
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
    except Exception as e:
        print(f"\n程序运行出错: {str(e)}")
    finally:
        print("\n程序已退出")
if __name__ == "__main__":
            main()