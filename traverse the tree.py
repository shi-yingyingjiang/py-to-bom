import win32com.client
import re 
from openpyxl import Workbook

component_dict = {}

# def traverse_assembly():
#     swApp = win32com.client.Dispatch("SldWorks.Application")
#     swModel = swApp.ActiveDoc
#     swAssy = swModel

#     if swAssy is not None:
#         if swAssy.IsVirtual:
#             swModel = swAssy.GetSubAssembly
#             if swModel is not None:
#                 traverse_assembly(swModel)
#         else:
#             swModel = swAssy
#         components = swAssy.GetComponents(False)  # 获取装配体中的所有组件
        
#         for comp in components:
#             if comp is not None:
#                 # print(f"组件名称: {comp.Name2}")
#                 txt = comp.Name2
#                 if "Part" not in txt:
#                     name = re.sub(r"-\d+$", "", txt)
#                     component_dict[name] = component_dict.get(name, 0)+1
#     return component_dict


# import win32com.client

# def traverse_assembly(swAssy):
#     if swAssy is None:
#         print("请打开一个装配体")
#         return

#     components = swAssy.GetComponents(False)  # 获取装配体中的所有组件
#     for comp in components:
#         if comp is not None:
#             if comp.IsVirtual:  # 检查是否为子装配体
#                 # 如果是子装配体，递归遍历子装配体中的零件
#                 sub_assy = comp.GetSubAssembly()
#                 print(sub_assy)
                # if sub_assy is not None:
                #     traverse_assembly(sub_assy)  # 递归调用
            # else:
                # 如果是零件，输出信息
                # print(f"零件名称: {comp.Name2}")
                # print(f"零件路径: {comp.GetPathName()}")
                # print(f"是否压缩: {comp.IsSuppressed}")
                # print(f"是否隐藏: {comp.IsHidden(False)}")

# def main():
#     swApp = win32com.client.Dispatch("SldWorks.Application")
#     swModel = swApp.ActiveDoc
#     if swModel is not None and swModel.GetType == 2:  # 2 表示装配体
#         traverse_assembly(swModel)
#     else:
#         print("请打开一个装配体")

# # 调用主函数
# if __name__ == "__main__":
#     main()



# def traverse_assembly(swAssy, component_dict):
#     """
#     递归遍历装配体及其子装配体中的组件
#     :param swAssy: 当前装配体
#     :param component_dict: 存储组件名称和数量的字典
#     :return: component_dict
#     """
#     components = swAssy.GetComponents(False)

#     for comp in components:
#         if comp is not None:
#             # 判断是否为子装配体
#             swCompModel = comp.GetModelDoc2()
#             # if swCompModel is not None and swCompModel.GetTypeName2() == "AssemblyDocument":
#             if swCompModel and "Assembly" in swCompModel.GetTypeName2():
#                 # 是子装配体，递归遍历
#                 traverse_assembly(swCompModel, component_dict)
#             else:
#                 # 是零件，不是装配体
#                 txt = comp.Name2
#                 if "Part" not in txt:
#                     name = re.sub(r"-\d+$", "", txt)
#                     component_dict[name] = component_dict.get(name, 0) + 1
#     return component_dict


# swApp = win32com.client.Dispatch("SldWorks.Application")
# swModel = swApp.ActiveDoc

# if swModel is not None:
#     component_dict = {}
#     traverse_assembly(swModel, component_dict)
#     print(component_dict)
# 调用函数
# traverse_assembly()

# components_dict = traverse_assembly()
# print(components_dict)


# def judgment(components_dict):
#     processeds_part_dict = {}
#     fasteners_dict = {}

#     components_dict = components_dict
#     for component in components_dict:
#         if component[:7].isdigit():
#             processeds_part_dict[component] = components_dict[component]
#         else:
#             fasteners_dict[component] = components_dict[component]
#     return processeds_part_dict, fasteners_dict


# the_parts = judgment(traverse_assembly())


# def arrangement(dict):
#     sorted_keys = sorted(dict.keys())
#     sorted_dict = {key: dict[key] for key in sorted_keys}
#     return sorted_dict


# wb = Workbook()
# ws = wb.create_sheet("机加工件")



import win32com.client

def get_assembly_component_types():
    part_list = []
    c_Assemblies_list = []
    try:
        # 启动或连接到SolidWorks应用程序
        sw_app = win32com.client.Dispatch("SldWorks.Application")
    except Exception as e:
        print("无法启动SolidWorks应用程序。请确保SolidWorks已安装并运行。")
        print(e)
        return

    # 获取当前活动文档
    sw_model = sw_app.ActiveDoc
    if not sw_model:
        print("没有打开的文档。")
        return

    # 检查是否为装配体文档
    if sw_model.GetType != 2:  # swDocumentTypes_e.swDocASSEMBLY
        print("当前文档不是装配体。")
        return

    # 获取装配体中的所有组件
    components = sw_model.GetComponents(False)
    if not components:
        print("装配体中没有组件。")
        return

    # 遍历组件并获取类型
    for comp in components:
        sw_comp = comp
        comp_model = sw_comp.GetModelDoc2
        if comp_model:
            doc_type = comp_model.GetType
            if doc_type == 1:  # swDocumentTypes_e.swDocPART
                print(f"组件名称: {sw_comp.Name}, 类型: 零件")
                part_list.append(sw_comp.Name)
            elif doc_type == 2:  # swDocumentTypes_e.swDocASSEMBLY
                print(f"组件名称: {sw_comp.Name}, 类型: 子装配体")
                c_Assemblies_list.append(sw_comp.Name)
            else:
                print(f"组件名称: {sw_comp.Name}, 类型: 未知")
        else:
            print(f"组件名称: {sw_comp.Name}, 类型: 无法获取模型文档")
    
    return part_list, c_Assemblies_list


# get_assembly_component_types()


part_list= get_assembly_component_types()
c_Assemblies_list = get_assembly_component_types()
def remove_prefixes(strings, prefixes):
    result = []
    for s in strings:
        if '/' in s:
            parts = s.split('/', 1)  # 只分割一次，最多分成两部分
            if len(parts) > 1 and parts[0] in prefixes:
                result.append(parts[1])
            else:
                result.append(s)
        else:
            result.append(s)
    return result

# 示例


processed_list = remove_prefixes(part_list,c_Assemblies_list)
# print(processed_list)
# 输出: ['banana', 'orange/grape', 'lemon', 'kiwi/lime', 'peach']


def classify_by_substrings(strings, keywords):
    Machined_parts = []
    Standard = []
    for s in strings:
        if any(keyword in s for keyword in keywords):
            Standard.append(s)
        else:
            Machined_parts.append(s)
    return Machined_parts,Standard

# 示例

the_Standard = ["apple", "pear"]

Machined_parts, Standard= classify_by_substrings(processed_list, the_Standard)

def the_number(thelist):
    number_dict = {}
    for item in thelist:
        number_dict[item] = number_dict.get(item, 0) + 1
        return number_dict


