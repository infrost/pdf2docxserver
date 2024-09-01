import socket
import json
import threading
from pdf2docx import Converter
from pathlib import Path
from tkinter import Tk, filedialog
from concurrent.futures import ThreadPoolExecutor

def parse_data(data):
    try:
        parsed_data = json.loads(data)
        source_path = parsed_data.get('source_path', None)
        output_path = parsed_data.get('output_path', None)
        arg = parsed_data.get('arg', None)
        return source_path, output_path, arg
    except json.JSONDecodeError:
        return None, None, None


def pdf_convert(source_path, output_path, multi_process):
    pdf_file = Path(source_path)
    docx_file = Path(output_path)

    cv = Converter(pdf_file)
    if multi_process:
        print("多进程处理已启用")
        cv.convert(docx_file, multi_processing=True)
    else:
        cv.convert(docx_file)
    cv.close()
    print(f"PDF转换完成: {output_path}")


def handle_client(client_socket, addr, executor):
    print(f"Connection established with {addr}")

    try:
        data = client_socket.recv(1024).decode('utf-8')
        print(f"Received from client: {data}")

        source_path, output_path, arg = parse_data(data)
        
        if source_path and output_path:
            # 将PDF转换任务提交给线程池处理，并等待其完成
            future = executor.submit(pdf_convert, source_path, output_path, arg)
            future.result()  # 阻塞直到转换完成

            response = "PDF conversion completed successfully."
        else:
            response = "Invalid data received from client."

        client_socket.send(response.encode('utf-8'))

    except Exception as e:
        print(f"Error handling client {addr}: {e}")
        response = f"Error: {str(e)}"
        client_socket.send(response.encode('utf-8'))
    finally:
        client_socket.close()
        print(f"Connection with {addr} closed")


def start_server(connect_port, executor):
    server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server_socket.bind(('localhost', connect_port))
    server_socket.listen(5)
    print(f"Server listening on port {connect_port}...")

    try:
        while True:
            client_socket, addr = server_socket.accept()
            # 使用线程池处理客户端请求
            executor.submit(handle_client, client_socket, addr, executor)
    except Exception as e:
        print(f"Server error: {e}")
    finally:
        server_socket.close()

if __name__ == "__main__":
    connect_port = 9999
    print("正在启动PDF文档转换服务\n更多用法请参见https://github.com/infrost/pdf2docxserver")

    # 创建线程池用于管理任务
    with ThreadPoolExecutor(max_workers=5) as executor:
        # 将服务器启动放到一个独立的线程中
        server_thread = threading.Thread(target=start_server, args=(connect_port, executor))
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
            executor.submit(pdf_convert, file_path, output_path, True)
