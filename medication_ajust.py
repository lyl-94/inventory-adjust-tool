import pandas as pd
from decimal import Decimal, getcontext
from openpyxl import load_workbook
from flask import Flask, request, send_file, render_template, make_response
import pandas as pd
import io

app = Flask(__name__)


@app.before_request
def set_global_precision():
    # 设置全局精度
    getcontext().prec = 10


@app.route('/upload', methods=['POST'])
def upload_file():
    """
    上传文件并调整库存的函数。
    
    此函数会接收上传的Excel文件和目标总价参数，根据库存调整算法对库存进行调整，
    然后返回调整后的库存Excel文件。
    
    函数主要流程：
    1. 检查请求中是否包含必要的'file'和'target'参数。
    2. 从请求中获取上传的Excel文件和目标总价。
    3. 读取Excel文件中的当前库存信息。
    4. 调用库存调整函数对库存进行调整。
    5. 如果库存调整成功，则计算扣减计划并更新Excel文件。
    6. 返回更新后的Excel文件供用户下载。
    
    返回值：
    如果成功，返回更新后的Excel文件；如果失败，返回包含错误信息的响应。
    """
    if 'file' not in request.files or 'target' not in request.form:
        response = make_response("缺少参数！", 400)
        response.mimetype = "text/plain"
        return response
    
    file = request.files['file']
    target = Decimal(request.form['target'])
    
    print("目标总价：" + str(target))

    # 读取excel
    cur_inventory = read_inventory_excel(file)
    cur_inventory = [(index, price, quantity) for index, (quantity, price) in enumerate(cur_inventory)]

    # 开始调整库存
    result = adjust_inventory(cur_inventory, target)

    if not result:
        response = make_response("库存调整失败！", 400)
        response.mimetype = "text/plain"
        return response

    # 计算最终的扣减计划
    plan = [(index, price, quantity, float(Decimal(price) * Decimal(quantity)), deduction) for index, price, quantity, deduction in result]

    # 创建excel
    output_file = update_inventory_excel(file, plan)
    
    if not output_file:
        response = make_response("库存调整失败！", 400)
        response.mimetype = "text/plain"
        return response
    
    return send_file(output_file, attachment_filename='modified.xlsx', as_attachment=True)


@app.route('/')
def index():
    """
    渲染并返回主页模板。
    
    该函数用于加载并返回'index.html'页面模板。
    """
    return render_template('index.html')


def adjust_inventory(inventory: list, target_total: Decimal) -> list:
    """
    调整库存至目标总价。
    
    该函数接收库存列表和目标总价作为输入，尝试通过减少库存数量来达到或接近目标总价。
    如果成功，将返回一个包含调整后库存和相应扣减数量的列表，列表中的元素为(索引, 价格, 调整后数量, 扣减数量)。
    如果无法调整到目标总价，将返回None。
    
    注意：
    1. 函数假设inventory列表中的每个元素都是一个三元组(index, price, quantity)。
    2. 如果需要减少的金额小于0（即目标总价高于当前总价），函数将直接返回None。
    3. 函数会先对库存按照数量进行降序排序，以便优先扣减数量较多的库存。
    4. 函数依赖于一个名为find_available_deduction_plan的外部函数来查找可用的扣减计划。
    """
    
    # 计算当前总价
    cur_total = sum(Decimal(price) * Decimal(quantity) for _, price, quantity in inventory)
    print("当前总价：" + str(cur_total))

    # 计算需要减少的金额
    deduction_needed = cur_total - Decimal(target_total)
    print("需要减少的金额：" + str(deduction_needed))

    # 如果需要减少的金额小于0，则返回None
    if deduction_needed < 0:
        return None

    # 按照库存数量从大到小排序
    inventory = sorted(inventory, key=lambda x: x[2], reverse=True)

    # 用于记录扣减计划
    deduction_plan = [0] * len(inventory)

    # 查找可用的扣减计划
    if (find_available_deduction_plan(deduction_plan, inventory, 0, deduction_needed)):
        print("库存调整成功！")

        # 把扣减计划合并到库存列表中
        result = [(index, price, quantity - deduction, deduction) for (index, price, quantity), deduction in zip(inventory, deduction_plan)]
        # 按照索引排序
        result = sorted(result, key=lambda x: x[0])
        return result
        
    else:
        print("无法调整库存。") 
        return None


def find_available_deduction_plan(deduction_plan:list, inventory: list, start_index: int, total_deduction: Decimal) -> bool:
    """
    递归查找可行的扣减计划
    
    通过递归方式，在给定的库存中查找是否存在一种扣减计划，使得总扣减值满足要求。
    
    Args:
        deduction_plan: 扣减计划列表，用于记录每种库存的扣减数量。
        inventory: 库存列表，每个元素为(库存标识, 单个库存扣减值, 库存数量)。
        start_index: 当前处理的库存索引。
        total_deduction: 还需要扣减的总值。
    
    Returns:
        bool: 如果找到可行的扣减计划，则返回True，否则返回False。
    
    注意：
        该函数会修改deduction_plan列表，记录找到的可行扣减计划。
    """
    if total_deduction == 0:
        return True
    elif start_index >= len(inventory):
        return False
    
    max_count = min(total_deduction // Decimal(inventory[start_index][1]), Decimal(inventory[start_index][2]))
    for count in range(int(max_count), -1, -1):
        # 将当前扣减值设置为count
        deduction_plan[start_index] = count
        # 将后面的扣减值设置为0
        deduction_plan[start_index+1:] = [0] * (len(deduction_plan) - start_index - 1)
        next_deduction = Decimal(total_deduction) - count * Decimal(inventory[start_index][1])
        if find_available_deduction_plan(deduction_plan, inventory, start_index + 1, next_deduction):
            return True
    return False  


def read_inventory_excel(file_path):
    """
    读取库存Excel文件并提取指定数据
    
    该函数用于从指定的Excel文件中读取库存数据，筛选并提取出符合条件的数据。
    函数首先会读取整个Excel文件，并从第二行开始处理数据（忽略标题行）。
    接着，函数会筛选出第一列值为数字的行，即确保这些行是有效的数据行。
    然后，函数会提取出第六列（数量）和第四列（价格）的数据。
    最后，函数会将提取的数量和价格数据合并为一个二元数组，并返回该数组。

    Args:
        file_path : Excel文件路径或FileStorage对象
    
    Returns:
        result (list): 一个包含(数量, 价格)二元组的列表。
    """
    df = pd.read_excel(file_path)

    # 从第二行开始读取
    df = df.iloc[0:]

    # 筛选第一列值为数字的行
    df = df[df.iloc[:, 0].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x))]
    
    # 提取第六列和第四列的数据
    quantities = df.iloc[:, 5].values
    price = df.iloc[:, 3].values

    # 将提取的数据合并为一个二元数组
    result = list(zip(quantities, price))

    return result


def update_inventory_excel(file, data):
    """
    更新库存Excel文件。
    
    根据提供的数据更新库存Excel文件，并返回更新后的Excel文件内容。
    
    具体实现步骤包括：
    1. 读取原始的Excel文件内容。
    2. 根据提供的数据长度和原始Excel文件的长度，确定处理的数据范围。
    3. 在Excel文件中增加“扣减库存数量”和“扣减库存金额”两列，并根据提供的数据填充这两列的内容。
    4. 根据提供的数据更新Excel文件中的“药品单价”和“药品库存数量”两列的内容。
    5. 使用openpyxl加载更新后的Excel文件，并在文件中添加“库存金额合计”和“扣减库存金额合计”的SUM公式。
    6. 返回更新后的Excel文件内容。
    
    Returns:
        io.BytesIO: 更新后的Excel文件内容。
    """
    # 读取Excel文件
    df = pd.read_excel(file, engine='xlrd')

    # 根据data的长度和df的长度进行处理
    min_length = min(len(data), len(df))
    
    # 增加一列 扣减库存数量
    deduction_list = [row[4] for row in data]
    deduction_list.extend([None] * (len(df)-min_length))
    df[f'扣减库存数量'] = deduction_list

    # 增加一列 扣减库存金额
    deduction_list = [row[4] * row[1] for row in data]
    deduction_list.extend([None] * (len(df)-min_length))
    df[f'扣减库存金额'] = deduction_list

    # 将二元数组的第一列药品单价写回Excel的第6列（索引为5）
    df.iloc[:min_length, 5] = [row[2] for row in data]

    # 将二元数组的第二列药品库存数量写回Excel的第7列（索引为6）
    df.iloc[:min_length, 6] = [row[3] for row in data]
    
    # 临时存储到内存
    temp = io.BytesIO()
    writer = pd.ExcelWriter(temp, engine='openpyxl')
    df.to_excel(writer, index=False)
    writer.save()
    temp.seek(0)

   
    # 使用openpyxl加载新的Excel文件
    workbook = load_workbook(temp)
    sheet = workbook.active

    # 库存金额合计
    # 合计行写入SUM公式
    sheet[f'G{min_length + 4}'] = f'=SUM(G2:G{min_length + 3})'

    # 扣减库存金额合计
    # 合计行写入SUM公式
    sheet[f'K{min_length + 4}'] = f'=SUM(K2:K{min_length + 3})'

    
    # 保存更改到内存中的最终Excel文件
    output_file = io.BytesIO()
    workbook.save(output_file)
    output_file.seek(0)
    return output_file


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=9000)