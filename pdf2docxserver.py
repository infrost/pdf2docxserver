import socket
import json
import threading
from pdf2docx import Converter
from pathlib import Path
from tkinter import Tk, filedialog

# 定义解析数据的函数
def parse_data(data):
    try:
        parsed_data = json.loads(data)
        source_path = parsed_data.get('source_path', None)
        output_path = parsed_data.get('output_path', None)
        arg = parsed_data.get('arg', None)
        return source_path, output_path, arg
    except json.JSONDecodeError:
        return None, None, None

# 定义处理客户端请求的函数
def handle_client(client_socket, addr):
    print(f"Connection established with {addr}")

    try:
        data = client_socket.recv(1024).decode('utf-8')
        print(f"Received from client: {data}")

        source_path, output_path, arg = parse_data(data)
               
        pdf_convert(Path(source_path), Path(output_path), arg)
        response = f"Server received convert request..."
        client_socket.send(response.encode('utf-8'))

    except Exception as e:
        print(f"Error handling client {addr}: {e}")
    finally:
        client_socket.close()
        print(f"Success: Connection with {addr} closed")

# 定义监听连接的函数
def start_server(connect_port):
    server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server_socket.bind(('localhost', connect_port))
    server_socket.listen(5)
    print(f"Server listening on port {connect_port}...")

    while True:
        client_socket, addr = server_socket.accept()

        # 创建一个新线程来处理客户端请求
        thread = threading.Thread(target=handle_client, args=(client_socket, addr))
        thread.start()

def pdf_convert(source_path, output_path, arg):
    pdf_file = source_path
    docx_file = output_path

    cv = Converter(pdf_file)
    if arg:
        print("多进程处理已启用")
        cv.convert(docx_file, multi_processing=True)      # all pages by default
    else:
        cv.convert(docx_file)
    cv.close()

if __name__ == "__main__":
    connect_port = 9999
    print("正在启动PDF文档转换服务\n更多用法请参见https://github.com/infrost/pdf2docxserver")

    # 将服务器启动放到一个独立的线程中
    server_thread = threading.Thread(target=start_server, args=(connect_port,))
    server_thread.daemon = True
    server_thread.start()
    print("服务已启动...\n你也可以按下Enter键手动打开一个PDF文档并将其转化为Word文档")
    while True:
        input("")
        # 创建Tkinter窗口但不显示
        root = Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(
            title='选择一个PDF文件', 
            filetypes=[
                ('PDF', '*.pdf'),
                ('All Files', '*.*')
                ]
            )
        
        if not file_path:
            print("没有选择文件。")
            continue
        output_path = Path(file_path).with_suffix('.docx')  # 默认输出路径
        thread = threading.Thread(target=pdf_convert, args=(file_path, output_path, False))
        thread.start()
