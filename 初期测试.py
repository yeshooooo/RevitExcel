
import pandas as pd

raw_data_path = r"H:\work\2022.8\盐坝高速成果表.xls"

# #############################################################导入数据#############################################################
# 给水
data_raw_J = pd.read_excel(raw_data_path, sheet_name="J".strip(), header=0, skiprows=2,
                           usecols=["点号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                    "Y", "地面", "管顶", "管内底", "埋深m"])
# print(data_raw_J)

# 污水
data_raw_W = pd.read_excel(raw_data_path, sheet_name=("W".replace("\\n", "")), header=0, skiprows=2,
                           usecols=["点号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                    "Y", "地面", "管顶", "管内底", "埋深m"])

# 雨水
data_raw_Y = pd.read_excel(raw_data_path, sheet_name=("Y".replace("\\n", "")), header=0, skiprows=2,
                           usecols=["点号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                    "Y", "地面", "管顶", "管内底", "埋深m"])

# 路灯

data_raw_LD = pd.read_excel(raw_data_path, sheet_name=("LD".replace("\\n", "")), header=0, skiprows=2,
                            usecols=["点号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                     "Y", "地面", "管顶", "管内底", "埋深m"])
# 电力
data_raw_L = pd.read_excel(raw_data_path, sheet_name=("L".replace("\\n", "")), header=0, skiprows=2,
                           usecols=["点号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                    "Y", "地面", "管顶", "管内底", "埋深m"])

# 信号
data_raw_XH = pd.read_excel(raw_data_path, sheet_name=("XH".replace("\\n", "")), header=0, skiprows=2,
                            usecols=["点号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                     "Y", "地面", "管顶", "管内底", "埋深m"])

# 燃气
data_raw_R = pd.read_excel(raw_data_path, sheet_name=("R".replace("\\n", "")), header=0, skiprows=2,
                           usecols=["点号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                    "Y", "地面", "管顶", "管内底", "埋深m"])

# 电信
data_raw_D = pd.read_excel(raw_data_path, sheet_name=("D".replace("\\n", "")), header=0, skiprows=2,
                           usecols=["点号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                    "Y", "地面", "管顶", "管内底", "埋深m"])

#data_raw = data_raw_J.append(data_raw_W).append(data_raw_Y).append(data_raw_LD).append(data_raw_L).append(data_raw_XH).append(data_raw_R).append(data_raw_D)
data_raw = pd.concat([data_raw_J, data_raw_W, data_raw_Y, data_raw_LD, data_raw_L, data_raw_XH, data_raw_R, data_raw_D],
                     ignore_index=True)
# print(data_raw.columns)
# print(data_raw)


# 整体思路
# 第一列非空为节点,添加到新行第一个数
# 全局寻找与第一列相连的点组成此段管线的起点终点还有属性

# print(type(data_raw))
###########################################################填充起点为空的键,防止后续因遍历出现复杂操作########################################
# 顺序遍历,如果为空,则赋值为 index -1

# new_point_data_raw = data_raw["点号"].fillna(value="kong")
# # print(data_raw["点号"])
# print(new_point_data_raw)
# data_raw = pd.join(new_point_data_raw,data_raw[])

# for index,data in data_raw.iterrows():
#
#     # if  isinstance(data["点号"],float) :
#     #     # data["点号"] = data_raw[index]
#     #     # print(data["点号"])
#     #     print()
#     print(index)
# for i in range(0, len(data_raw)):
#     # print(data_raw.iloc[i]["点号"], ":",type(data_raw.iloc[i]["点号"]),"\n")
#     # if isinstance(data_raw.iloc[i]["点号"],float):
#
#     print("--------------------------------end1---------------------------")
#     if pd.isnull(data_raw.iloc[i]["点号"]):
#         data_raw.iloc[i]["点号"] = data_raw.iloc[i - 1]["点号"]
#
#         print(data_raw.iloc[i]["点号"])

###### 为避免后续因数据类型产生麻烦,将所有为nan的赋值为 "kong" 字符串作为标记 ###################################################################
data = data_raw.fillna(value="kong")
# print(data)

####################### 总数据
for i in range(0,len(data)):
    if data.iloc[i]["点号"] == "kong" :
        # 这里竟然是引用传递..
        # 使用 DafaFrameming.loc[行名, 列名] = 值 的方式去赋值, 而不是使用DataFrame[][]的形式去赋值.不然会报错
        data.loc[i,"点号"] = data.loc[i - 1,"点号"]
        # print(data.iloc[i]["点号"])
# print(data)

######## 给水
dataJ = data_raw_J.fillna(value="kong")
for i in range(0,len(dataJ)):
    if dataJ.iloc[i]["点号"] == "kong" :
        dataJ.loc[i,"点号"] = dataJ.loc[i - 1,"点号"]
# print(dataJ)


###### 污水
dataW = data_raw_W.fillna(value="kong")
for i in range(0,len(dataW)):
    if dataW.iloc[i]["点号"] == "kong" :
        dataW.loc[i,"点号"] = dataW.loc[i - 1,"点号"]
# print(dataW)

###### 雨水
dataY = data_raw_Y.fillna(value="kong")
for i in range(0,len(dataY)):
    if dataY.iloc[i]["点号"] == "kong" :
        dataY.loc[i,"点号"] = dataY.loc[i - 1,"点号"]
# print(dataY)

## 路灯

dataLD = data_raw_LD.fillna(value="kong")
for i in range(0,len(dataLD)):
    if dataLD.iloc[i]["点号"] == "kong" :
        dataLD.loc[i,"点号"] = dataLD.loc[i - 1,"点号"]
# print(dataLD)

# 电力
dataL = data_raw_L.fillna(value="kong")
for i in range(0,len(dataL)):
    if dataL.iloc[i]["点号"] == "kong" :
        dataL.loc[i,"点号"] = dataL.loc[i - 1,"点号"]
# print(data_L)

# 交通信号
dataXH = data_raw_XH.fillna(value="kong")
for i in range(0,len(dataXH)):
    if dataXH.iloc[i]["点号"] == "kong" :
        dataXH.loc[i,"点号"] = dataXH.loc[i - 1,"点号"]
# print(dataXH)

# 燃气

dataR = data_raw_R.fillna(value="kong")
for i in range(0,len(dataR)):
    if dataR.iloc[i]["点号"] == "kong" :
        dataR.loc[i,"点号"] = dataR.loc[i - 1,"点号"]
# print(dataR)

# 电信
dataD = data_raw_D.fillna(value="kong")
for i in range(0,len(dataD)):
    if dataD.iloc[i]["点号"] == "kong" :
        dataD.loc[i,"点号"] = dataD.loc[i - 1,"点号"]
# print(dataD)


# print(type(data['点号']),":",data['点号'])
#########################################################导出翻模表格#####################################################################



# 如果点号不为空,则证明有这个点

#         构造新的点数组,通过连接点号字段创建起点终点,然后遍历键为连接点号,一直往下找,直到点号不为空停止
# 圆管跟方管分开导出

# 给水
# 连接点号属性用于简化去重
excelJoutR = pd.DataFrame(columns=["点号","连接点号","颜色","长","宽","x1","y1","z1","x2","y2","z2"])
excelJoutC = pd.DataFrame(columns=["点号","连接点号","颜色","直径","x1","y1","z1","x2","y2","z2"])
excelJCW = pd.DataFrame(columns=["点号","高度","类型"])
for i in range(0,len(dataJ)):
    startPoint = dataJ.loc[i,"点号"]
    linkPoint = dataJ.loc[i,"连接点号"]


#### 注意,会存在管径或断面为kong的情况,此节点可能为检修井或其他,此情况不能遗漏
    # 经判断这里管径或断面参数全是字符串
    # print(type(dataJ.loc[i, "管径或断面"]))


    # # 矩形管线
    # 存在管径大于1000的情况,所以字符串长度大于4
    if (len(dataJ.loc[i, "管径或断面"]) > 4) and (dataJ.loc[i, "管径或断面"] != "kong"):

        ## 设置标记用于确定是否添加,每轮重置为True
        isNew = True
        size = dataJ.loc[i, "管径或断面"].split("X")
        size = [float(s) for s in size]
        print(size)
        # 求管道中心点z1,单位转为mm,因为有的表管顶为空,所以,中心点为地面*1000-埋深m*1000-宽/2
        z1 = dataJ.loc[i, "地面"] * 1000 - dataJ.loc[i, "埋深m"] * 1000 - size[1]/2.0
## 因为终点可能不在当前表格中,所以终点坐标从全局表中寻找
        for j in range(0,len(data)):
            if data.loc[j,"点号"] == linkPoint:
                # z2
                z2 = data.loc[j,"地面"] * 1000 - data.loc[j,"埋深m"] * 1000 - size[1]

                # 遍历新数组,起点和终点都不相同的才添加
                for k in range(0,len(excelJoutR)):
                    if excelJoutR.loc[k,"点号"] == dataJ.loc[i,"连接点号"] and excelJoutR.loc[k,"连接点号"] == dataJ.loc[i,"点号"]:
                        # 如果相同则标记isNew为False
                        isNew = False
                # 如果标记为True,才是新点,进行添加
                if isNew:
                    # excelJoutR.append({"点号":startPoint,"连接点号":data.loc[j,"点号"],"颜色":"blue","长":size[0],"宽":size[1],
                    #                    "x1":dataJ.loc[i,"X"],"y1":dataJ.loc[i,"Y"],"z1":z1,
                    #                    "x2":data.loc[j,"X"],"y2":data.loc[j,"Y"],"z2":z2},ignore_index= True)

                    # append方法已废弃,改用concat
                    temp = pd.DataFrame({"点号":startPoint,"连接点号":data.loc[j,"点号"],"颜色":"blue","长":size[0],"宽":size[1],
                                       "x1":dataJ.loc[i,"X"],"y1":dataJ.loc[i,"Y"],"z1":z1,
                                       "x2":data.loc[j,"X"],"y2":data.loc[j,"Y"],"z2":z2},index=[0])
                    excelJoutR = pd.concat( [excelJoutR,temp],ignore_index= True)
                    # 找到第一个就停止,避免重复添加
                    break

    ## 圆形管线
    if (len(dataJ.loc[i, "管径或断面"]) <= 4) and (dataJ.loc[i, "管径或断面"] != "kong"):
        isNew = True
        size = float(dataJ.loc[i,"管径或断面"])
        # #这里经判断键为字符串,因此上面需转换为float
        # print(size)
        # print(type(size))
        z1 = dataJ.loc[i,"地面"] * 1000 - dataJ.loc[i,"埋深m"] * 1000 - size/2.0
        for j in range(0,len(data)):
            if data.loc[j,"点号"] == linkPoint:
                # z2
                z2 = data.loc[j,"地面"] * 1000 - data.loc[j,"埋深m"] * 1000 - size/2.0
                for k in range(0,len(excelJoutC)):
                    if excelJoutC.loc[k,"点号"] == dataJ.loc[i,"连接点号"] and excelJoutC.loc[k,"连接点号"] == dataJ.loc[i,"点号"]:
                        # 如果相同则标记isNew为False
                        isNew = False
                if isNew:

                    temp0 = pd.DataFrame({"点号":startPoint,"连接点号":data.loc[j,"点号"],"颜色":"blue","直径":size,
                                       "x1":dataJ.loc[i,"X"],"y1":dataJ.loc[i,"Y"],"z1":z1,
                                       "x2":data.loc[j,"X"],"y2":data.loc[j,"Y"],"z2":z2},index=[0])
                    excelJoutC = pd.concat( [excelJoutC,temp0],ignore_index= True)
                    # 找到第一个就停止,避免重复添加
                    break
    ## 只要有井字,就生成井
    ## 这里较单一,可以先添加再去重
    label = "井"
    if label in dataJ.loc[i,"附属物"]:
        ## 这里暂定为 地面高度,后面根据需求修改
        height = dataJ.loc[i,"地面"]
        temp1 = pd.DataFrame({"点号":dataJ.loc[i,"点号"],"高度":height,"类型": dataJ.loc[i,"附属物"]},index=[0])
        excelJCW = pd.concat([excelJCW,temp1],ignore_index=True)

    # 去重
    excelJCW = excelJCW.drop_duplicates()






## 导出管线
## https://blog.csdn.net/qq_35318838/article/details/104692846
## https://cloud.tencent.com/developer/ask/sof/1077612
# https://blog.csdn.net/dsy0221/article/details/120191310
# https://blog.csdn.net/u013247765/article/details/79050947
# 为避免麻烦,直接新建一个model.xlsx往里追加,否则追加模式会报错文件不存在
fileName = r"H:\work\2022.8\model.xlsx"
# if not os.path.exists(fileName):
#     os.system(r"touch {}".format(fileName))#调用系统命令行来创建文件

with pd.ExcelWriter(fileName,mode='a',engine="openpyxl",if_sheet_exists="replace") as writer:
    excelJoutR.to_excel(writer,sheet_name="矩形给水",index=False)
    excelJoutC.to_excel(writer,sheet_name="圆形给水",index=False)
    excelJCW.to_excel(writer,sheet_name="给水井",index=False)
