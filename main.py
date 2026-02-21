import logging
import os
import time
from datetime import datetime

import pandas as pd
from netmiko import ConnectHandler
from netmiko.exceptions import NetMikoTimeoutException, NetMikoAuthenticationException

# 配置日志
logging.basicConfig(
    filename='server.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger()


def load_devices_from_excel_conf(file_path, sheet_name=0):
    """
    从Excel文件加载设备信息
    :param file_path: Excel文件路径
    :param sheet_name: 工作表名称或索引
    :return: 设备列表字典
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # 检查必要列是否存在
        required_columns = ['ip', 'port', 'device_type', 'username', 'password']
        missing_cols = [col for col in required_columns if col not in df.columns]

        if missing_cols:
            logger.error(f"Excel配置文件中缺少必要列: {', '.join(missing_cols)}")
            return []

        # 处理可选列
        if 'port' not in df.columns:
            df['port'] = 22  # 默认SSH端口

        if 'secret' not in df.columns:
            df['secret'] = None  # 默认无特权密码

        if 'description' not in df.columns:
            df['description'] = ''  # 默认无描述

        # 转换为设备字典列表
        devices = df.to_dict('records')
        logger.info(f"成功从 {file_path} 加载 {len(devices)} 台设备")
        return devices

    except Exception as e:
        logger.error(f"读取Excel文件失败: {str(e)}")
        return []


def get_switch_config(device, input_command=None):
    """
    连接到交换机并获取配置
    :param device: 设备连接参数字典
    :param input_command: 用户输入的ssh操作命令
    :return: 配置文本或None
    """
    try:
        start_time = time.time()

        # 复制设备字典以避免修改原始数据
        conn_params = device.copy()
        print(f"\n{conn_params}")

        # 移除Netmiko不需要的额外字段
        for key in ['description']:
            if key in conn_params:
                del conn_params[key]

        # 建立SSH连接
        with ConnectHandler(**conn_params) as conn:
            # 进入特权模式（如果提供了secret）
            if conn_params.get('secret'):
                conn.enable()

            # 获取配置
            command = input_command if input_command else get_config_command(conn_params['device_type'])
            logger.info(f"execute command: {command} ......")
            output = conn.send_command(command, delay_factor=2)

            duration = time.time() - start_time
            logger.info(f"成功获取 {device['ip']} 配置 ({duration:.2f}秒)")
            return command, output

    except NetMikoTimeoutException:
        print(f"\n连接超时: {device['ip']}")
        logger.error(f"连接超时: {device['ip']}")
    except NetMikoAuthenticationException:
        print(f"\n认证失败: {device['ip']}")
        logger.error(f"认证失败: {device['ip']}")
    except Exception as e:
        print(f"\n获取 {device['ip']} 配置失败: {str(e)}")
        logger.error(f"获取 {device['ip']} 配置失败: {str(e)}")

    return None, None


def get_config_command(device_type):
    """
    根据设备类型返回获取配置的命令
    """
    commands = {
        'cisco_ios': 'show running-config',
        'cisco_nxos': 'show running-config',
        'huawei': 'display current-configuration',
        'h3c': 'display current-configuration',
        'hp_comware': 'display current-configuration',
        'juniper_junos': 'show configuration',
        'arista_eos': 'show running-config',
        'fortinet': 'show full-configuration',
        'paloalto_panos': 'show configuration running',
    }
    return commands.get(device_type, 'show running-config')


def save_config_to_file(ip, command, config, output_dir):
    """
    保存配置到文件
    :param ip: 设备IP地址
    :param command: 命令
    :param config: 配置文本
    :param output_dir: 输出目录
    :return: 保存的文件路径
    """
    try:
        # 操作时间
        op_time = datetime.now()
        # 创建文件名
        filename = f"{ip}_{op_time.strftime('%Y%m%d')}.txt"
        file_path = os.path.join(output_dir, filename)

        content = f"# {op_time.strftime('%Y-%m-%d %H:%M:%S')} {command}"
        # 如果文件不存在则创建，存在则追加
        with open(file_path, 'a' if os.path.exists(file_path) else 'w', encoding='utf-8') as f:
            f.write(content + "\n")
            f.write(config + "\n")
            f.write("## --=========================================================--" + "\n")

        logger.info(f"===> 配置已保存: {filename}")
        return file_path

    except Exception as e:
        logger.error(f"保存配置失败: {str(e)}")
        return None


def process_connection(devices, output_dir, input_command=None):
    """
    处理设备连接
    :param input_command: 用户输入的ssh操作命令
    :param output_dir: 配置文件输出目录
    :param devices:
    :return: 处理成功数量
    """
    # 处理每台设备
    success_count = 0
    print("\n===> start process collection switch devices configurations...")
    print("-" * 50)

    for device in devices:
        ip = device['ip']
        desc = device.get('description', '')

        print(f"===> connect to devices: {ip} {desc}...", end='', flush=True)

        # 获取配置
        command, config = get_switch_config(device, input_command)

        if config:
            # 保存配置
            save_config_to_file(ip, command, config, output_dir)
            success_count += 1
            print("\n ✓ execute successful...")
        else:
            print("\n ✗ execute failure...")

    # 输出结果
    print("\n" + "=" * 50)

    return success_count


def print_log(success_count, devices, is_from_config_file=True):
    """
    打印日志
    :param success_count: 处理成功数量
    :param devices: 设备列表
    :param is_from_config_file: 是否来自配置文件
    :return:
    """
    print(f"数据采集完毕 : {success_count}/{len(devices)}（成功数量/任务总数）")
    print("日志文件: server.log")
    config_path = "config_storage/normal" if is_from_config_file else "config_storage/other"
    print(f"配置文件目录: {config_path}")
    print("\n" + "-*-" * 20)


def execute_other_command(devices):
    """
    执行其他命令（如ping、traceroute等）
    :param devices:
    :return:
    """
    while True:
        try:
            input_command = input("\n请输入操作指令（输入 'exit' 退出）：")
            if input_command == 'exit':
                logger.info("===> 用户退出程序")
                break

            # 确认是否继续
            proceed = input("\n是否继续采集数据? (y/n): ").lower()
            if proceed != 'y':
                print("操作已取消。")
                break

            print(f"===> 正在执行命令: {input_command} ...")

            # 创建配置文件目录
            other_dir = os.path.join("config_storage", "other")
            if not os.path.exists(other_dir):
                os.mkdir(other_dir)

            success_count = process_connection(devices, other_dir, input_command)
            print_log(success_count, devices, False)
        except KeyboardInterrupt:
            print("\n程序被中断")
            break


def main():
    # 加载设备列表
    devices = load_devices_from_excel_conf("switch_ssh_conf.xlsx")

    if not devices:
        print("===> 未加载到任何设备，请检查Excel配置文件[switch_ssh_conf.xlsx]格式和内容。")
        return

    print(f"\n===> 已加载 {len(devices)} 台设备:")
    for i, device in enumerate(devices, 1):
        desc = f" ({device.get('description', '')})" if device.get('description') else ''
        print(f"{i}. {device['ip']} [{device['device_type']}]{desc}")

    # 确认是否继续
    proceed = input("\n是否开始采集交换机配置? (y/n): ").lower()
    if proceed != 'y':
        print("操作已取消。")
        return

    # 创建配置文件目录
    normal_path = os.path.join("config_storage", "normal")
    if not os.path.exists(normal_path):
        os.mkdir(normal_path)

    success_count = process_connection(devices, normal_path)
    print_log(success_count, devices)

    # 执行其他命令
    execute_other_command(devices)


if __name__ == "__main__":
    main()
