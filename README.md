# DeeplxFile
基于pdf2word提供的封装好的PDF转Word文档的服务程序

------------------
点击[这里下载](https://github.com/infrost/DeeplxFile/releases)

## 功能介绍
### 这个程序是干什么的？
这个程序提供了一个最大限度保留原PDF并将其转为docx的服务程序，可以配合[deeplxfile](https://github.com/infrost/DeeplxFile)使用。
也提供了api调用方法，可以将其集成在你的python应用中

## 我该如何使用？
### Windows
提供了封装好的exe版本, 运行安装程序即可
![](/images/server.png)
![](/images/sample.png)
### MacOS

**即将支持**

### 从源代码运行
你也可以下载源代码，然后运行
```bash
python pdf2docxserver.py
```

### API接口

``` Python
import socket
import json

def send_message(message):
    client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    client_socket.connect(('localhost', 9999))

    # 发送消息到服务器
    client_socket.send(message.encode('utf-8'))
    print(f"Sent to server: {message}")

    # 接收服务器的响应
    response = client_socket.recv(1024).decode('utf-8')
    print(f"Received from server: {response}")

    client_socket.close()

if __name__ == "__main__":
    data = {
        "source_path": r" ", #这里是原PDF路径
        "output_path": r" ", #这里是需要生成Word的输出路径
        "arg": True #是否需要启用多线程
    }
    message = json.dumps(data)
    send_message(message)

```


## 版本说明

```bash
--------- V0.1.0--------------

```
