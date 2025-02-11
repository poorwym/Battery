import subprocess
import sys
import os

def check_and_install_packages():
    """检查并安装所需的包"""
    required_packages = {
        'pandas': 'pandas',
        'matplotlib': 'matplotlib',
        'numpy': 'numpy',
        'openpyxl': 'openpyxl',
        'xlrd': 'xlrd',
        'scipy': 'scipy',  # 添加scipy包用于统计分析
        'seaborn': 'seaborn'  # 添加seaborn包用于更美观的统计图表
    }
    
    def install_package(package_name):
        print(f"正在安装 {package_name}...")
        try:
            # 添加超时参数和升级pip的选项
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", package_name], 
                               timeout=300)
            print(f"{package_name} 安装成功！")
            return True
        except subprocess.CalledProcessError:
            print(f"{package_name} 安装失败，请手动安装。")
            return False
        except subprocess.TimeoutExpired:
            print(f"{package_name} 安装超时，请检查网络连接或手动安装。")
            return False

    all_installed = True
    for package, pip_name in required_packages.items():
        try:
            __import__(package)
            print(f"{package} 已安装")
        except ImportError:
            print(f"未找到 {package}，准备安装...")
            if not install_package(pip_name):
                all_installed = False
    
    if not all_installed:
        print("\n有些包安装失败，请手动安装后再运行程序。")
        sys.exit(1)
    
    print("\n所有必需的包都已安装！\n")

# 运行环境检查
if __name__ == "__main__":
    check_and_install_packages()
    
    # 在确认环境后导入所需的包
    import pandas as pd
    import matplotlib.pyplot as plt
    import os
    import numpy as np

def get_excel_files():
    """获取当前目录下的所有Excel文件"""
    # 获取当前脚本所在目录的绝对路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"当前搜索目录：{current_dir}")  # 添加调试信息
    
    # 修改文件扩展名匹配逻辑，使其大小写不敏感
    excel_files = [
        f for f in os.listdir(current_dir) 
        if f.lower().endswith(('.xlsx', '.xls'))
    ]
    print(f"找到的Excel文件：{excel_files}")  # 添加调试信息
    
    return excel_files

def select_file(excel_files):
    """让用户选择要处理的Excel文件"""
    print("在当前目录下找到以下Excel文件：")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {file}")
    
    while True:
        try:
            choice = int(input("\n请选择要处理的文件编号（输入数字）: "))
            if 1 <= choice <= len(excel_files):
                # 确保使用正确的文件名（包括大小写）
                selected_file = excel_files[choice-1]
                file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), selected_file)
                if os.path.exists(file_path):
                    return file_path
                else:
                    print(f"错误：无法找到文件 {file_path}")
                    return None
            print("无效的选择，请重新输入！")
        except ValueError:
            print("请输入有效的数字！")

def process_data(df):
    """处理数据：重命名列、删除不需要的列和空白列，删除含-1和FF(%)>100的数据行"""
    # 显示原始数据行数
    print(f"\n原始数据行数：{len(df)}")
    
    # 调试信息：显示所有列名
    print("\n当前数据包含以下列：")
    for i, column in enumerate(df.columns, 1):
        print(f"{i}. {column}")
    print("\n")
    
    # 重命名和删除指定列
    df = df.rename(columns={'电池编号': 'name'})
    df = df.rename(columns={'PCE(%)':'PCE(%).1'})
    df = df.rename(columns={'Etac(%)':'PCE(%)'})
    df = df.drop('扫描次数', axis=1)
    
    # 删除全为空值的列
    df = df.dropna(axis=1, how='all')
    
    # 调试信息：显示处理后的列名
    print("处理后的数据包含以下列：")
    for i, column in enumerate(df.columns, 1):
        print(f"{i}. {column}")
    print("\n")

    df_1 = df.copy()

    # 删除指定列
    df = df.drop(['PCE(%).1', 'FF(%).1', 'Jsc(mA/cm2).1', 'Voc(V).1'], axis=1)
    df_1 = df_1.drop(['PCE(%)', 'FF(%)', 'Jsc(mA/cm2)', 'Voc(V)'], axis=1)
    
    # 显示 df_1 的列名和行数
    print("df_1 删除列后的列名：")
    for i, column in enumerate(df_1.columns, 1):
        print(f"{i}. {column}")
    print(f"df_1 的行数：{len(df_1)}\n")
    
    df_1 = df_1.rename(columns={
        'PCE(%).1': 'PCE(%)',
        'FF(%).1': 'FF(%)',
        'Jsc(mA/cm2).1': 'Jsc(mA/cm2)',
        'Voc(V).1': 'Voc(V)'
    })
    
    # 显示 df_1 重命名后的列名
    print("df_1 重命名后的列名：")
    for i, column in enumerate(df_1.columns, 1):
        print(f"{i}. {column}")
    print("\n")

    # 合并数据框（垂直方向）
    df_final = pd.concat([df, df_1], axis=0)
    print(f"合并后的数据行数：{len(df_final)}")
    print("\n合并后的name列内容：")
    for i, name in enumerate(df_final['name'], 1):
        print(f"{i}. {name}")
    print("\n")
    
    # 删除含有-1的行
    df_final = df_final[~(df_final == -1).any(axis=1)]
    print(f"删除含-1的行后，剩余数据行数：{len(df_final)}")
    print("\n删除-1后的name列内容：")
    for i, name in enumerate(df_final['name'], 1):
        print(f"{i}. {name}")
    print("\n")
    
    # 删除FF(%)大于100的行
    df_final = df_final[df_final['FF(%)'] <= 100]
    print(f"删除FF(%)>100的行后，剩余数据行数：{len(df_final)}")
    print("\n删除FF(%)>100后的name列内容：")
    for i, name in enumerate(df_final['name'], 1):
        print(f"{i}. {name}")
    print("\n")
    
    # 对于相同name的行，只保留PCE(%)最大的一行
    print("\n处理重复name前的详细数据：")
    print("重复的name值及其对应的PCE(%)值：")
    duplicated_names = df_final['name'].duplicated(keep=False)
    if any(duplicated_names):
        dup_data = df_final[duplicated_names].sort_values('name')
        print(dup_data[['name', 'PCE(%)']])
        
    # 先对数据框进行排序，确保PCE(%)最大的行会被保留
    df_final = df_final.sort_values('PCE(%)', ascending=False)
    # 删除重复的name，保留第一次出现的（即PCE(%)最大的）
    df_final = df_final.drop_duplicates(subset='name', keep='first')
    
    print(f"\n保留每个name中PCE(%)最大值后，剩余数据行数：{len(df_final)}")
    print("\n最终的name列内容：")
    for i, name in enumerate(df_final['name'], 1):
        print(f"{i}. {name}")
    
    # 再次检查是否还有重复
    final_duplicates = df_final['name'].duplicated(keep=False)
    if any(final_duplicates):
        print("\n警告：仍然存在重复的name：")
        print(df_final[final_duplicates][['name', 'PCE(%)']])
    
    # 重置索引并按name排序
    df_final = df_final.sort_values('name').reset_index(drop=True)
    
    return df_final

def setup_chinese_display():
    """配置matplotlib中文显示"""
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False

def create_boxplot(df, title):
    """为不同组的数据创建并列箱型图"""
    # 创建分组函数
    def get_group(name):
        # 分割name字符串
        parts = name.split('-')[0]
        return parts
    
    # 添加组别列
    df['group'] = df['name'].apply(get_group)
    
    # 创建不包含name列的数据框副本用于绘图
    plot_df = df.drop('name', axis=1)
    
    # 获取所有组别
    groups = sorted(df['group'].unique())
    
    # 获取需要绘制的列名（不包括group列）
    columns = [col for col in plot_df.columns if col != 'group']
    
    # 为每个数据列创建一个子图
    for col in columns:
        plt.figure(figsize=(10, 6))
        
        # 准备每组的数据
        data = [plot_df[plot_df['group'] == group][col] for group in groups]
        
        # 创建箱型图
        bp = plt.boxplot(data, 
                        labels=groups,
                        showfliers=True,
                        whis=np.inf,
                        showmeans=True,
                        meanprops={"marker":"s",
                                 "markerfacecolor":"white", 
                                 "markeredgecolor":"black",
                                 "markersize":6})
        
        # 为每组添加散点
        for i, group_data in enumerate(data, 1):
            # 生成对应的x坐标（添加少许随机偏移使点不重叠）
            x = np.random.normal(i, 0.04, size=len(group_data))
            plt.scatter(x, group_data, alpha=0.5, s=30)
        
        plt.title(f'{title} - {col}', fontsize=12)
        plt.xlabel('组别', fontsize=10)
        plt.ylabel('数值', fontsize=10)
        plt.tight_layout()
        plt.show()
        
        # 打印每组的数据统计
        print(f"\n{col} 的数据统计：")
        for group in groups:
            group_df = df[df['group'] == group]
            print(f"\n组 {group}:")
            print(f"样本数量：{len(group_df)}")
            print("包含的样品：")
            for i, name in enumerate(group_df['name'], 1):
                print(f"{i}. {name}")
        print("\n")

def create_boxplot_from_excel():
    """主函数：整合所有功能"""
    # 配置中文显示
    setup_chinese_display()
    
    # 获取Excel文件列表
    excel_files = get_excel_files()
    if not excel_files:
        print("当前目录下没有找到Excel文件！")
        return
    
    # 用户选择文件
    selected_file = select_file(excel_files)
    
    try:
        # 读取并处理数据
        df = pd.read_excel(selected_file)
        df = process_data(df)
        
        # 创建箱型图
        create_boxplot(df, selected_file)
        
    except Exception as e:
        print(f"处理文件时出错：{str(e)}")

if __name__ == "__main__":
    create_boxplot_from_excel()