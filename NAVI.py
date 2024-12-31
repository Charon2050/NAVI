import sys, os, _winapi
import msvcrt
import time, re, ast
import csv, json
import subprocess
import threading
from openai import OpenAI,AuthenticationError,APIConnectionError,BadRequestError,NotFoundError


# 用于支持Windows系统的彩色输出
if os.name == "nt":
	os.system("")


# 定义记忆文件保存路径
memory_file_path = os.path.join(os.getenv('APPDATA'), 'NAVI', 'memory.csv')

# 定义Log文件保存路径
log_file_path = os.path.join(os.getenv('APPDATA'), 'NAVI', 'NAVI_Log.log')

# 定义Config文件保存路径
config_file_path = os.path.join(os.getenv('APPDATA'), 'NAVI', 'NAVI_Config.cfg')

default_config = {
    'base_url': 'https://api.deepseek.com/beta',
    'model': 'deepseek-chat',
    'api_key': '',
    'user_name': "USER",
    'max_round': 20,
    'max_try_times': 5,
    'tts_volume': 80,
    'skip_auth': False,
    'simple_shell_output': True,
    'hide_shell_output': False,
    'example_mode': True,
    'quiet_mode': False,
}

# 写入 config 文件
def set_config(var,value):

    # 若无 NAVI_Config.cfg 则创建一个，并以 BOM 格式写入默认值
    if not os.path.exists(config_file_path):
        if not os.path.exists(os.path.dirname(config_file_path)):
            os.makedirs(os.path.dirname(config_file_path))
        with open(config_file_path, 'w+', encoding='utf-8-sig') as f:
            f.write(str(default_config).replace('{','{\n').replace('}','\n}').replace(', ',',\n'))

    # 如果文件为空，就以 BOM 格式写入默认值
    with open(config_file_path, 'r+', encoding='utf-8-sig') as f:
        if f.read()=='':
            f.truncate(0); f.seek(0)
            f.write(str(default_config).replace('{','{\n').replace('}','\n}').replace(', ',',\n'))

    # 写入修改值
    with open(config_file_path, 'r+', encoding='utf-8-sig') as f:
        try:
            config = f.read()
            # 如果能正确解析为字典
            if isinstance(ast.literal_eval(config),dict):
                # 写入
                config = ast.literal_eval(config)
                config.update({var:value})
                f.truncate(0); f.seek(0)
                f.write(str(config).replace('{','{\n').replace('}','\n}').replace(', ',',\n'))
                return f'Successfully set {var} to {value}'
        # 如果未能正确解析为字典
        except NameError:
            pass
        print('警告：NAVI_Config.cfg 文件似乎格式有误，请查修或删除。本次未修改 NAVI_Config.cfg 文件。')
        return 'WARNING: Format of NAVI_Config.cfg seems incorrect. Failed to read content.'


# 读取 config 文件
def read_config(var):

    # 若无 NAVI_Config.cfg 则创建一个，并以 BOM 格式写入
    if not os.path.exists(config_file_path):
        if not os.path.exists(os.path.dirname(config_file_path)):
            os.makedirs(os.path.dirname(config_file_path))
        with open(config_file_path, 'w+', encoding='utf-8-sig') as f:
            f.write(str(default_config).replace('{','{\n').replace('}','\n}').replace(', ',',\n'))

    # 如果文件为空，就以 BOM 格式写入默认值
    with open(config_file_path, 'r+', encoding='utf-8-sig') as f:
        if f.read()=='':
            f.truncate(0); f.seek(0)
            f.write(str(default_config).replace('{','{\n').replace('}','\n}').replace(', ',',\n'))

    with open(config_file_path, 'r', encoding='utf-8-sig') as f:
        try:
            config = f.read()
            # 如果能正确解析为字典
            if isinstance(ast.literal_eval(config),dict):
                config = ast.literal_eval(config)
                if config.get(var) is not None:
                    return config[var]         # 字典里有的键，返回对应的值
                else:
                    return default_config[var] # 字典里没有的键，返回默认值
        # 如果未能正确解析为字典
        except:
            pass
        print('警告：NAVI_Config.cfg 文件似乎已损坏，请查修或删除。')
        return default_config[var]     # 返回默认值


# 识别代码块的语言，键全小写
shell_mode={
    "navi_shell":"NAVI_Shell",
    "powershell":"PowerShell",
    "cmd":"CMD",
    #"python":"Python",
}

# 已测试: DeepSeek, 阿里通义(推荐), 智谱GLM, 讯飞星火(不推荐)
# 理论可支持任何兼容 OpenAI 格式调用的模型。但推荐使用代码能力较强的模型，不推荐使用免费模型。
base_url = read_config('base_url')
model = read_config('model')
api_key = read_config('api_key')

api_key_verified = False

# 定义用户名
user_name = read_config('user_name')

# 定义最长历史轮数
max_round = read_config('max_round')

# 初始化一轮对话中执行代码次数的计数器
code_try = 0

# 定义一轮对话中执行代码的最多次数
max_try_times = read_config('max_try_times')

# 语音音量，0-100
tts_volume = read_config('tts_volume')

# 这个 messages 就是历史记录，注意是不包括 system prompt 的
messages = []

# 用于储存后台进程的二维列表，每个元素都是一个有3个元素的列表，running_processes[x][0]为后台进程，running_processes[x][1]为此进程开始的时间，running_processes[x][2]为输出和错误
running_processes = []

# 等待用户输入时开启，用户输入后关闭，供check_completed_processes()判断自己是否仍需运行
waiting_input = False

# 预置示例对话，引导一些比较蠢的模型正确输出
def example_messages():

    default_example_messages = [
                                {"role": "user", "content": "请帮我打开文件资源管理器。"},
                                {"role": "assistant", "content": "好的，正在打开。\n\n```powershell\nStart-Process explorer\n```"},
                                {"role": "user", "content": "```SystemMessage\nShell compeleted with no error and no output.\n```"},
                                {"role": "assistant", "content": "已打开文件资源管理器。"},
                                {"role": "user", "content": "请帮我设置在今晚23:59关机"},
                                {"role": "assistant", "content": "> 分析：用户希望在指定时间关机，可以使用 `shutdown` 命令。\n>\n>为了计算倒计时时间，可以使用 `Get-Date` 计算时间差，然后用 `.TotalSeconds` 计算剩余秒数。\n\n正在设置...\n\n```powershell\n$shutdownTime = (Get-Date -Hour 23 -Minute 59 -Second 0)\n$timeSpan = $shutdownTime - (Get-Date)\n$seconds = [int]$timeSpan.TotalSeconds\nshutdown /s /t $seconds\n```"},
                                {"role": "user", "content": "```SystemMessage\nShell compeleted with no error and no output.\n```"},
                                {"role": "assistant", "content": "设置完成，你的电脑将在今晚23:59关机。"}
                              ]
    # 尝试打开 SampleMessages.json
    try:
        with open(running_path + 'SampleMessages.json', 'r', encoding='utf-8-sig') as f:
            # 尝试读取消息列表
            try:
                messages = json.load(f)
            # 如果读取失败，就创建一个空列表
            except json.JSONDecodeError:
                messages = default_example_messages
                print("SampleMessages.json 似乎已损坏，请尝试删除或修复。")
    # 如果不存在 SampleMessages.json，就创建一个
    except FileNotFoundError:
        messages = default_example_messages
        with open(running_path + 'SampleMessages.json', 'w') as f:
            json.dump(messages, f)
    # 返回消息列表
    return messages

# 测试API可用性，可用返回空字符串，不可用返回错误提示
def api_test(base_url, model, api_key):
    try:
        OpenAI(api_key=api_key, base_url=base_url).chat.completions.create(
            model=model,
            temperature=0.1,
            max_tokens=1,
            stream=False,
            messages=[{'role': 'user', 'content': '1'}]
        )
    except APIConnectionError:
        return '连接失败！请检查网络。如果使用了自定义 BaseURL，请检查 URL 是否设置正确。'
    except AuthenticationError:
        return 'API 验证失败！请检查 API 是否可用。如果使用了自定义 BaseURL，请检查 API 和 BaseURL 是否对应。'
    except BadRequestError:
        return '请求失败！如果使用了自定义模型，请检查模型名称是否正确（模型名称不一定是页面上宣传的商品名，请寻找模型编码 model 列表）'
    except NotFoundError:
        return '404 Not Found！如果使用了自定义 BaseURL，请检查 URL 是否设置正确。应当设置为 OpenAI 兼容的 URL。'
    except UnicodeEncodeError:
        return '输入错误！不支持中文或特殊字符'
    else:
        return ''
    
# 自动解码
def auto_decode(data):
    # 尝试列表
    encodings = ['UTF-8','GBK']
    # 判断是否是字节类型
    if type(data)==bytes:
        # 列出的所有编码类型逐个尝试
        for temp in encodings:
            try:
                return data.decode(temp)
            except UnicodeDecodeError:
                pass
        # 若全部失败则抛出异常
        raise UnicodeDecodeError(f'Auto decode failed when trying to decode with {' '.join(encodings)}')
    # 若是列表，则递归到单个值
    elif type(data) in [list,tuple]:
        temp=[]
        for i in range(len(data)):
            temp.append(auto_decode(data[i]))
        return temp
    # 若是字符串则直接返回其自身，增强鲁棒性
    elif type(data)==str:
        return data
    # 没有就返回空字符串
    elif data is None:
        return ''
    # 其他类型抛出异常
    else:
        raise TypeError(f'auto_decode() can only decode Data or a list of Data, not {type(data)}')


# 循环检测是否有完成的后台进程
def check_completed_processes():

    # 只在等待用户输入时检测
    while waiting_input:
        completed_processes=[]
        i=0

        #遍历全部后台进程
        while i < len(running_processes):

            # 此处参考 https://blog.csdn.net/KiteRunner/article/details/129848482
            # 若进程仍在运行
            if running_processes[i][0].poll() is None:
                # 获取handle
                handle = msvcrt.get_osfhandle(running_processes[i][0].stdout.fileno())
                # 若有输出
                try: # 下面这行执行时，有小概率报错「管道已结束」
                    if _winapi.PeekNamedPipe(handle, 0)[0] > 0:
                        # 读取输出
                        data= _winapi.ReadFile(handle, _winapi.PeekNamedPipe(handle, 0)[0])[0]
                        try:
                            # 尝试解码
                            running_processes[i][2] = running_processes[i][2] + auto_decode(data) # 输出是自带换行的
                        except UnicodeDecodeError:
                            # 否则报错
                            running_processes[i][2] = "Error: Processes finished, but failed to read the output. It may caused by incorrect file encoding / decoding."
                            # 已经编码错误了就别试了
                            running_processes[i][0].kill()
                except BrokenPipeError:
                    pass

            # 若进程已结束
            else:
                # 添加一行提示消息（暂不带代码块）
                try:
                    result = running_processes[i][2] + "".join(auto_decode(running_processes[i][0].communicate()))
                except:
                    result = "Error: Processes finished, but failed to read the output. It may caused by incorrect file encoding / decoding."

                # 缩减过长输出
                result = re.sub(r'(\s)\1{3,}', r'\1\1', result) # 删除连续3个以上的空白字符
                if len(result) > 2048:
                    result = result[:1024]+"\n......\n" + result[-1024:]

                # 添加到完成列表
                completed_processes.append(f"Info: A background program (started at {running_processes[i][1]}) just finished, with following result:\n{result}")
                # 释放进程
                running_processes[i][0].kill()
                running_processes.pop(i)
                continue
            i = i + 1 

        # 遍历后，如果发现了至少1个完成的进程
        if completed_processes:
            write_log('\n'.join(completed_processes))
            # 输出消息之前，先换一行，防止和 'USER: ' 输入提示放在同一行
            print('')
            # 输出系统消息
            if hide_shell_output :
                0
            elif simple_shell_output :
                print("\033[34m<<<<< "+completed_processes[0][:70]+"...\033[0m")
            else:
                print("\033[34m<<<<< "+"\n<<<<< ".join(completed_processes)+"\033[0m")
            # 添加到历史消息
            messages.append({
                "content": '```SystemMessage\n'+'\n'.join(completed_processes)+'\n```',
                "role": "user"
                })
            # 传给AI
            write_log("Sent "+str(messages))
            response = client.chat.completions.create(
                model=model,
                temperature=1,
                max_tokens=1000,
                stop=["\n```\n\n"],
                stream=False,
                messages=system_prompt_messages()+messages
            )
            response.choices[0].message.content = fix_response(response.choices[0].message.content)
            write_log("Received "+response.choices[0].message.content)

            # 结果加入历史记录
            messages.append({"role": "assistant", "content": response.choices[0].message.content})
            # 输出结果（日后要改成流式的）
            output_message(response.choices[0].message.content)
            # 如果 AI 继续写 shell，就run_shell()
            for i in response.choices[0].message.content.split("\n"):
                if i[:3] == "```" and i[3:].lower() in shell_mode:
                    write_log("LOGS: Detected shells, start to run it")
                    # 如果没有代码块结尾就补一个，防止代码块有头无尾
                    if messages[-1].get("content").count('```') % 2 == 1:
                        messages[-1]=dict(messages[-1], **{"content": messages[-1].get("content")+"\n```"})
                    run_shell()
                    break
            # 输出一个 'USER: ' 输入提示，提示用户依然处于等待输入状态
            print("\033[1;33m"+user_name+": "+"\033[0m",end='',flush=True) # 有些控制台必须刷新才能显示出不换行的消息
        time.sleep(2)


def now_time():
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())


def system_prompt_messages():

    # 从 SystemPrompt.md 中读取 SystemPrompt
    with open('SystemPrompt.md', 'r', encoding='utf-8-sig') as f:
        system_prompt = f"当前时间：{now_time()}\n\n用户昵称：`{user_name}`\n\n当前运行路径：`{running_path}\\`\n\n{f.read()}"

    # 若无memory.csv则创建一个
    if not os.path.exists(memory_file_path):
        if not os.path.exists(os.path.dirname(memory_file_path)):
            os.makedirs(os.path.dirname(memory_file_path))
        with open(memory_file_path, mode='w', encoding='utf-8-sig', newline='') as file:
            writer = csv.writer(file)
            writer.writerows([])
            return [{"role":"system","content":system_prompt+"\n\n看起来，这是你与这位用户的第一次见面。如果用户没有要求你做事，可以先收集并记录这台电脑的信息，用 powershell 查询一下 CPU、GPU、内存、硬盘分区和总容量、用户名等信息，然后用 NAVI_Shell 代码块记住这些信息。"}]+example_messages()*example_mode
    
    # 读取memory.csv并返回
    with open(memory_file_path, mode='r', encoding='utf-8-sig') as file:
        memory_list = [item for sublist in list(csv.reader(file)) for item in sublist]
        if len(memory_list) == 0: 
            return [{"role":"system","content":system_prompt+"\n\n看起来，这是你与这位用户的第一次见面。如果用户没有要求你做事，可以先收集并记录这台电脑的信息，用 powershell 查询一下 CPU、GPU、内存、硬盘分区和总容量、用户名等信息，然后用 NAVI_Shell 代码块记住这些信息。"}]+example_messages()*example_mode
        else:
            return [{"role":"system","content":system_prompt+"\n\n目前 memory.csv 中已知的记忆信息：\n\n"+"\n".join(memory_list)}]+example_messages()*example_mode


def write_log(log):

    # 若无.log文件则创建一个
    if not os.path.exists(log_file_path):
        if not os.path.exists(os.path.dirname(log_file_path)):
            os.makedirs(os.path.dirname(log_file_path))
        with open(log_file_path, 'w', encoding='utf-8-sig') as f:
            pass  # 创建文件时不写入任何内容

    # 在不读取文件的情况下直接写入一行log
    with open(log_file_path, 'a', encoding='utf-8-sig') as f:
        f.write('[' + now_time() + "] " + log + '\n')

def voice_speek(content,voice=['Chinese','Japanese','English']):

    def find_voice_name(key_words):
        shells = [
            '$voices = (New-Object -ComObject SAPI.SpVoice).GetVoices()',
            '$voices | ForEach-Object { $_.GetDescription() }'
        ]
        voices = subprocess.Popen(["powershell", "-Command", "\n".join(shells)],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,creationflags=subprocess.CREATE_NO_WINDOW).communicate()[0]
        
        if type(key_words) == str:
            key_words = [key_words]

        for temp in key_words:
            for i in voices.split('\n'):
                if temp.lower() in i.lower():
                    return i[:i.find(' - ')]

    shells = [
        'Add-Type -AssemblyName System.speech',
        '$tts=New-Object System.Speech.Synthesis.SpeechSynthesizer',
        '$tts.SelectVoice("'+find_voice_name(voice)+'")',
        '$tts.Volume='+str(tts_volume),
        '$tts.Speak("'+content+'")',
        '#tts.Dispose()']
    subprocess.Popen(["powershell", "-Command", "\n".join(shells)],stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,creationflags=subprocess.CREATE_NO_WINDOW)


def fix_response(content):

    # 补丁：有些模型不遵守stop词，所以手动删掉第二个```之后的所有内容
    if content.count('```') > 1:
        content = content[:content.find('```',content.find('```')+1)]

    # 防止模型自己输出 SystemMessage
    if content.count('```SystemMessage') > 0:
        content = content[:content.find('```SystemMessage')]
        if content == '':
            content = '\n'

    # 防止模型输出没有标注语言的代码块
    if content.find('```') == content.find('```\n') and content.find('```')!=-1:
        # 把代码块开头和结尾删掉
        content = content[:content.find('```')] + content[content.find('```')+3:]
        content = content[:content.find('```')] + content[content.find('```')+3:]

    return content


def navi_shell(shell):
    global tts_volume
    global user_name
    # 若无memory.csv则创建一个
    if not os.path.exists(memory_file_path):
        write_log('No memory.csv found, creating...')
        with open(memory_file_path, mode='w', encoding='utf-8-sig', newline='') as file:
            writer = csv.writer(file)
            writer.writerows([])
        write_log('Created memory.csv')

    if shell[:9]=='remember ':
        # 读取memory.csv到data
        with open(memory_file_path, mode='r', encoding='utf-8-sig') as file:
            reader = csv.reader(file)
            data = list(reader)
        # 检查是否已有此记忆
        if [shell[9:]] in data:
            return f'NAVI_Shell Error: Already remembered "{shell[9:]}"'
        # 增加一行记忆
        data.append([shell[9:]])
        # 储存
        with open(memory_file_path, mode='w', encoding='utf-8-sig', newline='') as file:
            writer = csv.writer(file)
            writer.writerows(data)
            write_log(f'Added to memory.csv: {data[-1][0]}')
            return f'remembered {shell[9:]}'

    elif shell[:7]=="forget ":
        # 读取memory.csv到data
        with open(memory_file_path, mode='r', encoding='utf-8-sig') as file:
            reader = csv.reader(file)
            data = list(reader)
        # 删除此记忆
        i = 0
        temp = f'NAVI_Shell Error: No such memory "{shell[7:]}"'
        while i < len(data):
            if data[i] == [shell[7:]]:
                data.pop(i)
                temp = "forgot "+shell[7:]
                write_log(f'Deleted from memory.csv: {shell[7:]}')
                continue
            i = i + 1
        # 储存
        with open(memory_file_path, mode='w', encoding='utf-8-sig', newline='') as file:
            writer = csv.writer(file)
            writer.writerows(data)
            return temp

    elif shell[:13]=="check_process":
        temp = []
        for i in running_processes:
            # 不管状态，全部都告诉AI仍在运行。因为运行完成后 check_completed_processes() 会自动汇报，如果这里也汇报，AI 就会收到两条完成消息，导致编造其中一条的结果
            # 如果没有输出
            if i[2].replace('\n','').replace(' ','') == '':
                temp.append(f'INFO: A process started at {i[1]} is still running.')
            # 如果已有输出
            else:
                # 删除连续3个以上的空白字符
                i[2] = re.sub(r'(\s)\1{3,}', r'\1\1', i[2])
                # 缩减过长输出
                if len(i[2]) > 1024:
                    i[2] = i[2][:512]+"\n......\n" + i[2][-512:]
                temp.append(f'INFO: A process started at {i[1]} is still running with following content: \n' + i[2])
        return "\n".join(temp)+'\nINFO: No other process running. Do not repeatedly run this command.'
    
    elif shell[:6]=="volume":
        if shell=="volume":
            return tts_volume
        if shell[7:].isdigit():
            if int(shell[7:])>=0 and int(shell[7:])<=100:
                tts_volume = int(shell[7:])
                set_config('tts_volume',tts_volume)
                return f'Volume set to {str(tts_volume)}'
        return 'Error: Volume only support a Integer in 1-100.'
    
    elif shell[:9]=="user_name":
        if shell=="user_name":
            return user_name
        if shell[10:] and not ('\n' in shell[10:]):
            user_name = shell[10:]
            set_config('user_name',user_name)
            return f'user_name set to {user_name}'
        return 'Error: user_name only support a one-line string without quotation mark.'
        
    
    elif shell[:2]=='//' or shell=='':
        return ''
        
    else:
        return f'NAVI_Shell Error: No such command "{shell.split(" ")[0]}". Is it a PowerShell or CMD command?'


def user_input(message=""):

    # 清空此前尝试执行命令的次数
    global code_try
    code_try = 0

    # 用户输入
    while message == "":
        message=input("\033[1;33m"+user_name+": "+"\033[0m")
    write_log('User Input: '+message)

    global waiting_input
    waiting_input = False

    # 写入历史记录
    messages.append({"role": "user", "content": message})

    # 请求并获取回复（注意stop词是"\n```\n\n"，而不是"\n```"，否则会在代码块开头处被stop）
    write_log("Sent "+str(messages))
    response = client.chat.completions.create(
        model=model,
        temperature=1,
        max_tokens=1000,
        stop=["\n```\n\n"],
        stream=False,
        messages=system_prompt_messages()+messages
    )
    response.choices[0].message.content = fix_response(response.choices[0].message.content)
    write_log("Received "+response.choices[0].message.content)

    # 写入历史记录
    messages.append({"role": "assistant", "content": response.choices[0].message.content})
    # 输出结果（日后要改成流式的）
    output_message(response.choices[0].message.content)
    
    # 如果有可执行的 Shell 就执行
    for i in messages[-1].get("content").split("\n"):
        if i[:3] == "```" and i[3:].lower() in shell_mode:
            # 如果没有代码块结尾就补一个，防止代码块有头无尾
            if messages[-1].get("content").count('```') % 2 == 1:
                messages[-1]=dict(messages[-1], **{"content": messages[-1].get("content")+"\n```"})
            run_shell()
            break

    # 清理过长的记录
    if len(messages) > max_round*2 :
        messages.pop(0)
        messages.pop(0)
        write_log('History messages too long, cleared. Current history messages: '+str(messages))


def run_shell():

    # 把消息分行储存
    shells = messages[-1].get("content").split("\n")

    # 如果最后一行是代码块结尾，就删掉
    if shells[-1] == "```": shells.pop()

    # 从后往前找，删掉代码块开头之前（不含）的行
    for i in range(len(shells)-1,-1,-1):
        if re.search(r"```\w+",shells[i]):
            shells = shells[i:]
            break
    
    # 确定shell的语言
    run_mode=shell_mode.get(shells[0][3:].lower(),shells[0][3:]) # 若找不到，则返回原文

    # 删掉代码块开头，只留下命令行
    shells.pop(0)

    # 执行结果将储存至 result
    result = []
    # 输出要执行的命令
    if hide_shell_output :
        0
    elif simple_shell_output :
        print("\033[34m>>>>> "+shells[0][:70]+"...\033[0m")
    else:
        print("\033[34m>>>>> "+"\n>>>>> ".join(shells)+"\033[0m")

    # 是否显示窗口
    creation_flag = subprocess.CREATE_NEW_CONSOLE
    if hide_shell_output or simple_shell_output: creation_flag = subprocess.CREATE_NO_WINDOW

    # 执行 PowerShell
    if run_mode == "PowerShell":
        result = f"Info: The program started at {now_time()} has been running for too long and has been redirected to the background."

        # ------------------------------------------------
        # BUG: 运行多行命令时，如果结果类型不同，则会只输出第一种结果所属类型的结果。这会导致AI看不到结果时编造SystemMessage。测试用例：
        # "Get-CimInstance -ClassName Win32_LogicalDisk -Filter \"DriveType=3\" | Select-Object DeviceID,FreeSpace,Size",
        # "Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object FreePhysicalMemory,TotalVisibleMemorySize",
        # "$memory = Get-CimInstance Win32_OperatingSystem | Select-Object TotalVisibleMemorySize, FreePhysicalMemory",
        # "$memoryUsage = [math]::Round((($memory.TotalVisibleMemorySize - $memory.FreePhysicalMemory) / $memory.TotalVisibleMemorySize) * 100, 2)",
        # 临时补丁：以Get-CimInstance开头的命令加上" | Write-Host"。一个很烂的补丁，日后必须修
        # ------------------------------------------------
        for i in range(len(shells)):
            if shells[i][:15] == 'Get-CimInstance':
                shells[i] = shells[i] + " | Write-Host"
        # ------------------------------------------------
        #                                    必须修！↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
        running_processes.append([subprocess.Popen(["powershell", "-Command", "\n".join(shells)], stdout=subprocess.PIPE,stderr=subprocess.PIPE,creationflags=creation_flag),now_time(),''])
        write_log('Shell running: powershell -Command ' + '\n'.join(shells))

        # 等待4秒，超时后转入后台运行
        for i in range(10):
            time.sleep(0.4)
            # 若进程已结束
            if running_processes[-1][0].poll() is not None:
                try:
                    result = running_processes[-1][2] + "".join(auto_decode(running_processes[-1][0].communicate()))
                except UnicodeDecodeError:
                    result = "Error: Processes finished, but failed to read the output. It may caused by incorrect file encoding / decoding."
                running_processes[-1][0].kill()
                running_processes.pop()
                break

    # 执行 CMD
    elif run_mode == "CMD":
        result = f"Info: The program started at {now_time()} has been running for too long and has been redirected to the background."
        
        # CMD直接拼接多行命令易导致变量问题，所以改为创建脚本文件运行
        with open(os.path.join(os.getenv('TEMP'), 'cmd_script.bat'), "w") as f:
            # 在前面加个 @echo off 少输出一点，省点 Token
            f.write("@echo off\n"+"\n".join(shells))
        
        running_processes.append([subprocess.Popen(os.path.join(os.getenv('TEMP'), 'cmd_script.bat'),shell=True,stdout=subprocess.PIPE,stderr=subprocess.PIPE,creationflags=creation_flag),now_time(),''])
        write_log('Shell running: powershell -Command ' + '\n'.join(shells))

        # 等待4秒，超时后转入后台运行
        for i in range(10):
            time.sleep(0.4)

            # 若进程已结束
            if running_processes[-1][0].poll() is not None:
                try:
                    result = running_processes[-1][2] + "".join(auto_decode(running_processes[-1][0].communicate()))
                except:
                    result = "Error: Processes finished, but failed to read the output. It may caused by incorrect file encoding / decoding."
                running_processes.pop()
                break
            
    # 执行 NAVI_Shell
    elif run_mode == "NAVI_Shell":
        result = ""
        for i in shells:
            result = result + navi_shell(i) + "\n"
            write_log('Shell running: ' + i)
        result = result[:-1]

    elif run_mode == "SystemMessage":
        result = '''ERROR: NO PROGRAM FINISHED.
                WARNING: SYSTEM MESSAGE CODE BLOCKS CAN ONLY OUTPUT BY SYSTEM, YOU CAN NOT OUTPUT SYSTEM MESSAGE CODE. NEVER DO THIS AGAIN.'''

    # 其他情况报错
    else:
        result = f"Error: Can not run this language: {run_mode}. Only support PowerShell, CMD and NAVI_Shell."

    # 解释一下空的输出，防止 AI 以为失败了
    if result == '':
        result = 'Shell compeleted with no error and no output.'

    # 缩减过长结果并输出
    result = re.sub(r'(\s)\1{3,}', r'\1\1', result) # 删除连续3个以上的空白字符
    if len(result)>2048:
        result = result[:1024]+"\n......\n" + result[-1024:]
    if not (simple_shell_output or hide_shell_output):
        print("\033[34m<<<<< " + result + "\033[0m")
    write_log('Shell result: ' + result)

    # 将结果添加至历史记录
    messages.append({
        "content": f"```SystemMessage\n{result}\n```",
        "role": "user"
    }) 

    # 执行次数累加器，一轮对话内执行过多次代码强制打断
    global code_try
    code_try += 1
    write_log("Sent "+str(messages))
    if code_try <= max_try_times :
        # 带着新的历史记录传给AI
        response = client.chat.completions.create(
            model=model,
            temperature=1,
            max_tokens=1000,
            stop=["\n```\n\n"],
            stream=False,
            messages=system_prompt_messages()+messages
        )
    else:
        # 要求停止
        write_log("Warning: Runned too many times in one turn, requesting AI to stop...")
        response = client.chat.completions.create(
            model=model,
            temperature=1,
            max_tokens=100,
            stop=["\n```\n\n"],
            stream=False,
            messages=[{"role":"system","content":"你必须立刻简短的告诉用户，上述操作重复次数过多，被系统强制中断了。"}]+[messages[-1]]
        )
    response.choices[0].message.content = fix_response(response.choices[0].message.content)
    write_log("Received "+response.choices[0].message.content)

    # 结果加入历史记录
    #messages[-1].update({
    #    "content": messages[-1].get("content")+response.choices[0].message.content,
    #    "role": "assistant"
    #    })
    messages.append({"role": "assistant", "content": response.choices[0].message.content})
    # 输出结果（日后要改成流式的）
    output_message(response.choices[0].message.content)

    # 如果 AI 继续写可执行的 shell，就递归执行
    for i in response.choices[0].message.content.split("\n"):
        if i[:3] == "```" and i[3:].lower() in shell_mode:
            # 如果没有代码块结尾就补一个，防止代码块有头无尾
            if messages[-1].get("content").split("\n")[-1] != "```":
                messages[-1]=dict(messages[-1], **{"content": messages[-1].get("content")+"\n```"})
            run_shell()
            break


def output_message(message,no_new_line=False):

    voice_content = ''

    # 只输出可运行的 Shell 之前的内容
    for i in range(len(message.split("\n"))):
        if message.split("\n")[i][:3] == "```" and message.split("\n")[i][3:].lower() in shell_mode:
            message = "\n".join(message.split("\n")[:i])
            break

    # 逐行输出
    for i in message.split("\n"):
        if i != "" :    # 不输出空行
            if i[:2] == '> ':
                if not (hide_shell_output or simple_shell_output):
                    print(f"\033[34m>>>>> {i}\033[0m")
            else:
                print("\033[1;36mNAVI: " + "\033[0m"+i,end='\n'*(not(no_new_line))) # 如果no_new_line=True则不换行
                voice_content = voice_content + i + '；'

    # 播放语音
    if not quiet_mode:
        voice_speek(voice_content)


if __name__ == 'main' or True:
    
    write_log('')
    write_log('')
    write_log('----- New Program Start -----')
    # 处理参数
    running_path = os.getcwd() + '\\'
    sys.argv.pop(0)
    skip_auth = read_config('skip_auth')
    simple_shell_output = read_config('simple_shell_output')
    hide_shell_output = read_config('hide_shell_output')
    example_mode = read_config('example_mode')
    quiet_mode = read_config('quiet_mode')
    i = 0
    while i < len(sys.argv):
        if sys.argv[i].lower() in ["-k","-key","-apikey","-api-key","-api_key"]:
            api_key=sys.argv[i+1]
            sys.argv.pop(i)
            sys.argv.pop(i)
            write_log('api_key set to: '+api_key[:4]+'*'*(len(api_key)-8)+api_key[-4:])
            continue
        if sys.argv[i].lower() in ["-s","-skip"]:
            api_key_verified = True
            sys.argv.pop(i)
            write_log('api_key_verified set to: True')
            continue
        if sys.argv[i].lower() in ["-url","-baseurl","-base-url","-base_url"]:
            base_url=sys.argv[i+1]
            sys.argv.pop(i)
            sys.argv.pop(i)
            write_log(f'base_url set to: {base_url}')
            continue
        if sys.argv[i].lower() in ["-m","-model"]:
            model=sys.argv[i+1]
            sys.argv.pop(i)
            sys.argv.pop(i)
            write_log(f'model set to: {model}')
            continue
        if sys.argv[i].lower() in ["-noshell","-hideshell","-onlychat"]:
            # hide_shell_output 和 simple_shell_output 均开启时，hide_shell_output 优先生效
            hide_shell_output = True
            sys.argv.pop(i)
            write_log('hide_shell_output set to: True')
            continue
        if sys.argv[i].lower() in ["-simpleshell","-spsh","-lessshell","-lssh"]:
            simple_shell_output = True
            sys.argv.pop(i)
            write_log('simple_shell_output set to: True')
            continue
        if sys.argv[i].lower() in ["-example","-eg"]:
            example_mode=True
            sys.argv.pop(i)
            write_log('example_mode set to: True')
            continue
        if sys.argv[i].lower() in ["-quiet","-q",'-slience']:
            quiet_mode=True
            sys.argv.pop(i)
            write_log('quiet_mode set to: True')
            continue
        i = i + 1
    

    # 取得并验证APIKey可用性
    while not api_key_verified:
        if api_key == '':
            api_key = input('请输入 API KEY: ')
        test_result = api_test(base_url,model,api_key)
        if test_result == '':
            api_key_verified = True
            # 写入配置文件，但如果已设置过，就不覆盖
            if read_config('api_key') == '':
                set_config('api_key',api_key)
        else:
            print(test_result)
            api_key = ''

        
    # 定义 OpenAI 客户端
    client = OpenAI(api_key=api_key, base_url=base_url)
    
    # 如果带消息参数启动，就先执行一轮
    if " ".join(sys.argv):
        write_log('Start with user input: '+" ".join(sys.argv))
        user_input(" ".join(sys.argv))

    # 主循环
    while True:

        waiting_input = True
        check_thread = threading.Thread(target=check_completed_processes)
        check_thread.daemon = True
        check_thread.start()

        # 用户输入
        user_input()
