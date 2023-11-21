import markdown_to_json
import pandas as pd


def flatten_structure(key, value, result, current_level):
    """
    递归函数，用于展平嵌套的字典结构。
    """
    # 检查值是字典还是列表/单个元素
    if isinstance(value, dict):
        # 对于字典，递归调用此函数
        for sub_key, sub_value in value.items():
            flatten_structure(sub_key, sub_value, result, current_level + [key])
    elif isinstance(value, list):
        # 列表处理：对于列表中的每个元素，将当前层级的键和值添加到结果中
        for item in value:
            add_to_result(result, current_level + [key], item)
    else:
        # 单个元素处理：直接将键和值添加到结果中
        add_to_result(result, current_level + [key], value)


def add_to_result(result, keys, value):
    """
    将键和值添加到结果中的正确层级。
    """
    for i, key in enumerate(keys, start=1):
        result[f"h{i}"].append(key)
    result[f"h{len(keys) + 1}"].append(value)


def nest2flat(nested_dict: dict) -> dict:
    """
    将嵌套的字典结构转换为平面结构。
    """
    # 初始化结果字典，假设最多6层
    result = {f"h{i}": [] for i in range(1, 6 + 1)}

    # 遍历嵌套字典，并展平它
    for key, value in nested_dict.items():
        flatten_structure(key, value, result, [])

    return result


def main():
    with open("./input.md", "r", encoding="utf-8") as f:
        md_text = f.read()

    flat_dict = nest2flat(markdown_to_json.dictify(md_text))
    df = pd.DataFrame(
        {
            "h1": flat_dict["h1"],
            "h2": flat_dict["h2"],
            "h3": flat_dict["h3"],
            "h4": flat_dict["h4"],
            "h5": flat_dict["h5"],
        }
    )
    df.to_excel("./output.xlsx", index=False)


if __name__ == "__main__":
    main()
