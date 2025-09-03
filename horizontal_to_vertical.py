import re


def extract_specific_elements(original_list):
    lists_dict = {f'v{i}': [] for i in range(1, 13)}
    v1 = original_list[0::12]
    return {f'v{i+1}': original_list[i::12] for i in range(12)}

def is_all_empty_strings(lst):
    """
    检查列表是否仅包含空字符串('')
    
    参数:
        lst (list): 要检查的列表
        
    返回:
        bool: 如果列表中所有元素都是空字符串('')则返回True，否则返回False
    """
    return all(item == "" for item in lst)

def check_dict_empty_string_lists(data):
    """
    检查字典中所有值为列表的元素是否仅包含空字符串('')
    
    参数:
        dictionary (dict): 要检查的字典
        
    返回:
        dict: 返回一个新字典，包含原始字典中所有值为列表的键，
              对应的值为检查结果(True/False)
    """
 
    dictionary = {f'v{i+1}': data[i::12] for i in range(12)}


    result = {}
    for key, value in dictionary.items():
        if isinstance(value, list):
            result[key] = is_all_empty_strings(value)

    false_keys = []
    for key, value in dictionary.items():
        if isinstance(value, list):
            if not is_all_empty_strings(value):
                false_keys.append(value)
    return false_keys

# 示例使用
if __name__ == "__main__":
    sample_list = ['', 6095.0712890625, '', 10949.4111328125, '', 13456.37109375, '', 10878.4521484375, '', 17517.193359375, '', 13053.974609375, '', 5465.6845703125, '', 17635.2734375, '', 15073.353515625, '', 10057.65625, '', 11283.6748046875, '', 8547.728515625, '', 11088.775390625, '', 11817.9736328125, '', 11892.5107421875, '', 10945.794921875, '', 12643.5751953125, '', 8329.1591796875, '', 11968.908203125, '', 16271.1298828125, '', 14888.15234375, '', 9846.5322265625, '', 11430.6865234375, '', 7577.8505859375, '', 13510.0009765625, '', 18928.783203125, '', 17694.923828125, '', 13403.1982421875, '', 13009.365234375, '', 10708.13671875, '', 12178.0732421875, '', 18631.88671875, '', 14836.18359375, '', 13168.796875, '', 10924.7275390625, '', 10649.06640625, '', 19281.181640625, '', 17036.7265625, '', 14000.8232421875, '', 12497.0595703125, '', 15448.60546875, '', 10884.529296875, '', 16763.654296875, '', 14641.1708984375, '', 14437.9736328125, '', 16777.55859375, '', 15318.1748046875, '', 4522.3740234375]
  # 测试用列表
    result = extract_specific_elements(sample_list)
    lst = check_dict_empty_string_lists(result)
    print(lst)