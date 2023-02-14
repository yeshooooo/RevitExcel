import pandas as pd
# 原始数据格式化
def getData(raw_data_path,outputfile,isOffeset = True):
    '''
    :param raw_data_path: 原始数据路径
    :param outputfile: 偏移值写入excel第一个sheet,为了方便这里传入文件名后,后面的函数进行sheet追加模式
                        就不需要进行提前创建文件了
    :param isOffeset: 是否进行偏移,默认为True
    :return: 由datafrme组成的元组
    '''
    # 给水
    data_raw_J = pd.read_excel(raw_data_path, sheet_name="J".strip(), header=0, skiprows=2,
                               usecols=["管线点预编号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                        "Y", "地面", "管顶", "管底", "埋深"])
    # print(data_raw_J)

    # 污水
    data_raw_W = pd.read_excel(raw_data_path, sheet_name=("W".replace("\\n", "")), header=0, skiprows=2,
                               usecols=["管线点预编号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                        "Y", "地面", "管顶", "管底", "埋深"])

    # 雨水
    data_raw_Y = pd.read_excel(raw_data_path, sheet_name=("Y".replace("\\n", "")), header=0, skiprows=2,
                               usecols=["管线点预编号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                        "Y", "地面", "管顶", "管底", "埋深"])

    # 路灯

    data_raw_LD = pd.read_excel(raw_data_path, sheet_name=("LD".replace("\\n", "")), header=0, skiprows=2,
                                usecols=["管线点预编号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征",
                                         "附属物", "X",
                                         "Y", "地面", "管顶", "管底", "埋深"])
    # 电力
    data_raw_L = pd.read_excel(raw_data_path, sheet_name=("L".replace("\\n", "")), header=0, skiprows=2,
                               usecols=["管线点预编号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                        "Y", "地面", "管顶", "管底", "埋深"])

    # 交通信号
    data_raw_XH = pd.read_excel(raw_data_path, sheet_name=("XH".replace("\\n", "")), header=0, skiprows=2,
                                usecols=["管线点预编号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征",
                                         "附属物", "X",
                                         "Y", "地面", "管顶", "管底", "埋深"])

    # 燃气
    data_raw_R = pd.read_excel(raw_data_path, sheet_name=("R".replace("\\n", "")), header=0, skiprows=2,
                               usecols=["管线点预编号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                        "Y", "地面", "管顶", "管底", "埋深"])

    # 电力通讯
    data_raw_DT = pd.read_excel(raw_data_path, sheet_name=("DT".replace("\\n", "")), header=0, skiprows=2,
                               usecols=["管线点预编号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                        "Y", "地面", "管顶", "管底", "埋深"])

    # 电信
    data_raw_D = pd.read_excel(raw_data_path, sheet_name=("D".replace("\\n", "")), header=0, skiprows=2,
                               usecols=["管线点预编号", "连接点号", "埋设方式", "管线材料", "管径或断面", "特征", "附属物", "X",
                                        "Y", "地面", "管顶", "管底", "埋深"])

    # data_raw = data_raw_J.append(data_raw_W).append(data_raw_Y).append(data_raw_LD).append(data_raw_L).append(data_raw_XH).append(data_raw_R).append(data_raw_D)
    data_raw = pd.concat(
        [data_raw_J, data_raw_W, data_raw_Y, data_raw_LD, data_raw_L, data_raw_XH, data_raw_R, data_raw_DT,data_raw_D],
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
    for i in range(0, len(data)):
        if data.iloc[i]["管线点预编号"] == "kong":
            # 这里竟然是引用传递..
            # 使用 DafaFrameming.loc[行名, 列名] = 值 的方式去赋值, 而不是使用DataFrame[][]的形式去赋值.不然会报错
            data.loc[i, "管线点预编号"] = data.loc[i - 1, "管线点预编号"]
            # print(data.iloc[i]["点号"])
    # print(data)

    ######## 给水
    dataJ = data_raw_J.fillna(value="kong")
    for i in range(0, len(dataJ)):
        if dataJ.iloc[i]["管线点预编号"] == "kong":
            dataJ.loc[i, "管线点预编号"] = dataJ.loc[i - 1, "管线点预编号"]
    # print(dataJ)
    dataJ["管径或断面"] = dataJ["管径或断面"].astype("str")


    ###### 污水
    dataW = data_raw_W.fillna(value="kong")
    for i in range(0, len(dataW)):
        if dataW.iloc[i]["管线点预编号"] == "kong":
            dataW.loc[i, "管线点预编号"] = dataW.loc[i - 1, "管线点预编号"]
    # print(dataW)
    dataW["管径或断面"] = dataW["管径或断面"].astype("str")


    ###### 雨水
    dataY = data_raw_Y.fillna(value="kong")
    for i in range(0, len(dataY)):
        if dataY.iloc[i]["管线点预编号"] == "kong":
            dataY.loc[i, "管线点预编号"] = dataY.loc[i - 1, "管线点预编号"]
    # print(dataY)
    dataY["管径或断面"] = dataY["管径或断面"].astype("str")

    ## 路灯

    dataLD = data_raw_LD.fillna(value="kong")
    for i in range(0, len(dataLD)):
        if dataLD.iloc[i]["管线点预编号"] == "kong":
            dataLD.loc[i, "管线点预编号"] = dataLD.loc[i - 1, "管线点预编号"]
    # print(dataLD)
    dataLD["管径或断面"] = dataLD["管径或断面"].astype("str")

    # 电力
    dataL = data_raw_L.fillna(value="kong")
    for i in range(0, len(dataL)):
        if dataL.iloc[i]["管线点预编号"] == "kong":
            dataL.loc[i, "管线点预编号"] = dataL.loc[i - 1, "管线点预编号"]
    # print(data_L)
    dataL["管径或断面"] = dataL["管径或断面"].astype("str")

    # 交通信号
    dataXH = data_raw_XH.fillna(value="kong")
    for i in range(0, len(dataXH)):
        if dataXH.iloc[i]["管线点预编号"] == "kong":
            dataXH.loc[i, "管线点预编号"] = dataXH.loc[i - 1, "管线点预编号"]
    # print(dataXH)
    dataXH["管径或断面"] = dataXH["管径或断面"].astype("str")

    # 燃气

    dataR = data_raw_R.fillna(value="kong")
    for i in range(0, len(dataR)):
        if dataR.iloc[i]["管线点预编号"] == "kong":
            dataR.loc[i, "管线点预编号"] = dataR.loc[i - 1, "管线点预编号"]
    dataR["管径或断面"] = dataR["管径或断面"].astype("str")
    # print(dataR)
    # 电力通讯

    dataDT = data_raw_DT.fillna(value="kong")
    for i in range(0, len(dataDT)):
        if dataDT.iloc[i]["管线点预编号"] == "kong":
            dataDT.loc[i, "管线点预编号"] = dataDT.loc[i - 1, "管线点预编号"]
    dataDT["管径或断面"] = dataDT["管径或断面"].astype("str")
    # 电信
    dataD = data_raw_D.fillna(value="kong")
    for i in range(0, len(dataD)):
        if dataD.iloc[i]["管线点预编号"] == "kong":
            dataD.loc[i, "管线点预编号"] = dataD.loc[i - 1, "管线点预编号"]
    dataD["管径或断面"] = dataD["管径或断面"].astype("str")
    # print(dataD)
    if isOffeset:
        # 保持小数点后两位
        # middleX = data["X"].mean().round(2)
        # middleY = data["Y"].mean().round(2)
        middleX = (max(data["X"]) - min(data["X"]) ) /2
        middleY = (max(data["Y"]) - min(data["Y"]) ) /2

        # 偏移值描述,写入excel文件,在此描述的话,就不用提前建model.excel文件了
        offsetDesc = pd.DataFrame(columns=["是否进行偏移","偏移X","偏移Y"])
        temOffset = pd.DataFrame({"是否进行偏移":isOffeset,"偏移X": middleX,"偏移Y":middleY},index=[0])
        offsetDesc = pd.concat([offsetDesc,temOffset],ignore_index=True)
        offsetDesc.to_excel(outputfile,sheet_name="偏移描述",index=False)


        data["X"] = data["X"].sub(middleX)
        data["Y"] = data["Y"].sub(middleY)
        dataJ["X"] = dataJ["X"].sub(middleX)
        dataJ["Y"] = dataJ["Y"].sub(middleY)
        dataW["X"] = dataW["X"].sub(middleX)
        dataW["Y"] = dataW["Y"].sub(middleY)
        dataY["X"] = dataY["X"].sub(middleX)
        dataY["Y"] = dataY["Y"].sub(middleY)
        dataLD["X"] = dataLD["X"].sub(middleX)
        dataLD["Y"] = dataLD["Y"].sub(middleY)
        dataL["X"] = dataL["X"].sub(middleX)
        dataL["Y"] = dataL["Y"].sub(middleY)
        dataXH["X"] = dataXH["X"].sub(middleX)
        dataXH["Y"] = dataXH["Y"].sub(middleY)
        dataR["X"] = dataR["X"].sub(middleX)
        dataR["Y"] = dataR["Y"].sub(middleY)
        dataDT["X"] = dataDT["X"].sub(middleX)
        dataDT["Y"] = dataDT["Y"].sub(middleY)
        dataD["X"] = dataD["X"].sub(middleX)
        dataD["Y"] = dataD["Y"].sub(middleY)
        print("--------------------偏移sheet写入完成---------------------------------")


        return data, dataJ, dataW, dataY, dataLD, dataL, dataXH, dataR, dataDT,dataD
    else:
        offsetDesc = pd.DataFrame(columns=["是否进行偏移","偏移X","偏移Y"])
        temOffset = pd.DataFrame({"是否进行偏移":isOffeset,"偏移X": "不进行偏移","偏移Y":"不进行偏移"},index=[0])
        offsetDesc = pd.concat([offsetDesc,temOffset],ignore_index=True)
        offsetDesc.to_excel(outputfile,sheet_name="偏移描述",index=False)
        print("--------------------偏移sheet写入完成---------------------------------")
        return data, dataJ, dataW, dataY, dataLD, dataL, dataXH, dataR, dataDT,dataD

def exportExcel(dataAll,output_file_name):
    '''

    :param dataAll: getData函数返回的由datafrme组成的元组
    :param output_file_name: 输出文件名,为节省不必要的工作量,提前创建一个xlsx的文件,往文件里追加表格,如果文件不存在会报错
    :return: 无
    '''
    data = dataAll[0]

    for ii in range(1, len(dataAll)):
    # 单跑电力
    # for ii in range(5, 6):
        # 确定颜色
        # https://www.sojson.com/rgb.html
        # 根据getData函数导出, 1 为给水, 2 为污水, 3 为雨水, 4为路灯, 5为电力, 6 为交通信号, 7为燃气, 8为电力通讯, 9 电信



        #         构造新的点数组,通过连接点号字段创建起点终点,然后遍历键为连接点号,一直往下找,直到点号不为空停止
        # 圆管跟方管分开导出


        # 连接点号属性用于简化去重
        excelJoutR = pd.DataFrame(columns=["管线点预编号", "连接点号", "长", "宽", "x1", "y1", "z1", "x2", "y2", "z2"])
        excelJoutC = pd.DataFrame(columns=["管线点预编号", "连接点号", "直径", "x1", "y1", "z1", "x2", "y2", "z2"])


        # 拐点--> 特征为 "拐点"或"变径"均为 双通，只提取附属物为"kong"的点

        # 给水附属物 只有两种 检修井 和消火栓
        excelJXJ = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 检修井
        excelXHS = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 消火栓
        excelJCJ = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 检查井
        excelHFC = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 化粪池
        excelYB = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 雨篦
        excelYLK = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 预留口
        excelDD = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 地灯
        excelCD = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 出地
        excelLD = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 路灯
        excelPDX = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 配电箱
        excelJXX = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 接线箱
        excelKZG = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 控制柜
        excelXHD = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 信号灯
        excelSKJ = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 手孔井
        excelFST = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 电信发射塔
        excelRKJ = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 人孔井
        excelSXT = pd.DataFrame(columns=["管线点预编号", "地面", "类型","x","y","z"]) # 摄像头
        for i in range(0, len(dataAll[ii])):
            startPoint = dataAll[ii].loc[i, "管线点预编号"]
            linkPoint = dataAll[ii].loc[i, "连接点号"]

            #### 注意,会存在管径或断面为kong的情况,此节点可能为检修井或其他,此情况不能遗漏
            # 经判断这里管径或断面参数全是字符串
            # print(type(dataAll[ii].loc[i, "管径或断面"]))

            # # 矩形管线
            # # 存在管径大于1000的情况,所以字符串长度大于4
            # if (len(dataAll[ii].loc[i, "管径或断面"]) > 4) and (dataAll[ii].loc[i, "管径或断面"] != "kong"):

            # 存在float 100.0 大于4的情况,改为是否包含X
            if ("X" in dataAll[ii].loc[i, "管径或断面"]) and (dataAll[ii].loc[i, "管径或断面"] != "kong"):
                ## 设置标记用于确定是否添加,每轮重置为True
                isNew = True
                size = dataAll[ii].loc[i, "管径或断面"].split("X")
                size = [float(s) for s in size]
                # print(size)
                # print("矩形")
                # 求管道中心点z1
                # z1 = dataAll[ii].loc[i, "地面"] * 1000 - dataAll[ii].loc[i, "埋深m"] * 1000 - size[1] / 2.0
                # 因为管顶和管底很多数据是空的，所以这里中心点： 地面 - 埋深 + 半径 除以 2，单位是m
                z1 = float(dataAll[ii].loc[i, "地面"] - dataAll[ii].loc[i, "埋深"]) + size[1] / 2.0 / 1000.0
                ## 因为终点可能不在当前表格中,所以终点坐标从全局表中寻找
                for j in range(0, len(data)):
                    if data.loc[j, "管线点预编号"] == linkPoint:
                        # z2
                        # z2 = data.loc[j, "管顶"] * 1000 - data.loc[j, "埋深m"] * 1000 - size[1]
                        z2 = float(data.loc[j, "地面"] - data.loc[j, "埋深"]) + size[1] / 2.0 / 1000.0

                        # 遍历新数组,起点和终点都不相同的才添加
                        for k in range(0, len(excelJoutR)):
                            if excelJoutR.loc[k, "管线点预编号"] == dataAll[ii].loc[i, "连接点号"] and excelJoutR.loc[k, "连接点号"] == \
                                    dataAll[ii].loc[i, "管线点预编号"]:
                                # 如果相同则标记isNew为False
                                isNew = False
                        # 如果标记为True,才是新点,进行添加
                        if isNew:

                            temp = pd.DataFrame(
                                {"管线点预编号": startPoint, "连接点号": data.loc[j, "管线点预编号"], "长": size[0],
                                 "宽": size[1],
                                 "x1": dataAll[ii].loc[i, "X"], "y1": dataAll[ii].loc[i, "Y"], "z1": z1,
                                 "x2": data.loc[j, "X"], "y2": data.loc[j, "Y"], "z2": z2}, index=[0])
                            excelJoutR = pd.concat([excelJoutR, temp], ignore_index=True)
                            # 找到第一个就停止,避免重复添加
                            break

                ## 添加附属物
                label = "检修井"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelJXJ = pd.concat([excelJXJ, temp1], ignore_index=True)


                label = "消火栓"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelXHS = pd.concat([excelXHS, temp1], ignore_index=True)

                label = "检查井"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelJCJ = pd.concat([excelJCJ, temp1], ignore_index=True)
                label = "化粪池"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelHFC = pd.concat([excelHFC, temp1], ignore_index=True)

                label = "雨篦"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelYB = pd.concat([excelYB, temp1], ignore_index=True)

                label = "预留口"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelYLK = pd.concat([excelYLK, temp1], ignore_index=True)


                label = "地灯"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelDD = pd.concat([excelDD, temp1], ignore_index=True)

                label = "出地"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelCD = pd.concat([excelCD, temp1], ignore_index=True)



                label = "路灯"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelLD = pd.concat([excelLD, temp1], ignore_index=True)


                label = "配电箱"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelPDX = pd.concat([excelPDX, temp1], ignore_index=True)


                label = "接线箱"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelJXX = pd.concat([excelJXX, temp1], ignore_index=True)


                label = "控制柜"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelKZG = pd.concat([excelKZG, temp1], ignore_index=True)

                label = "信号灯"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelXHD = pd.concat([excelXHD, temp1], ignore_index=True)

                label = "手孔井"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelSKJ = pd.concat([excelSKJ, temp1], ignore_index=True)

                label = "电信发射塔"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelFST = pd.concat([excelFST, temp1], ignore_index=True)

                label = "人孔井"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelRKJ = pd.concat([excelRKJ, temp1], ignore_index=True)


                label = "摄像头"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelSXT = pd.concat([excelSXT, temp1], ignore_index=True)


            ## 圆形管线
            # if (len(dataAll[ii].loc[i, "管径或断面"]) <= 4) and (dataAll[ii].loc[i, "管径或断面"] != "kong"):
            if (not (("X" in dataAll[ii].loc[i, "管径或断面"]))) and (dataAll[ii].loc[i, "管径或断面"] != "kong"):
                isNew = True
                size = float(dataAll[ii].loc[i, "管径或断面"])
                # #这里经判断键为字符串,因此上面需转换为float
                # print(size)
                # print("圆形")
                # print(type(size))
                # z1 = dataAll[ii].loc[i, "地面"] * 1000 - dataAll[ii].loc[i, "埋深m"] * 1000 - size / 2.0
                # z1 = float(dataAll[ii].loc[i, "管顶"]) - size / 2.0 / 1000.0
                z1 = float(dataAll[ii].loc[i, "地面"] - dataAll[ii].loc[i, "埋深"]) + size / 2.0 / 1000.0
                for j in range(0, len(data)):
                    if data.loc[j, "管线点预编号"] == linkPoint:
                        # z2
                        # z2 = data.loc[j, "地面"] * 1000 - data.loc[j, "埋深m"] * 1000 - size / 2.0

                        z2 = float(data.loc[j, "地面"] - data.loc[j, "埋深"]) + size / 2.0 / 1000.0
                        for k in range(0, len(excelJoutC)):
                            if excelJoutC.loc[k, "管线点预编号"] == dataAll[ii].loc[i, "连接点号"] and excelJoutC.loc[k, "连接点号"] == \
                                    dataAll[ii].loc[i, "管线点预编号"]:
                                # 如果相同则标记isNew为False
                                isNew = False
                        if isNew:
                            temp0 = pd.DataFrame(
                                {"管线点预编号": startPoint, "连接点号": data.loc[j, "管线点预编号"], "直径": size,
                                 "x1": dataAll[ii].loc[i, "X"], "y1": dataAll[ii].loc[i, "Y"], "z1": z1,
                                 "x2": data.loc[j, "X"], "y2": data.loc[j, "Y"], "z2": z2}, index=[0])
                            excelJoutC = pd.concat([excelJoutC, temp0], ignore_index=True)
                            # 找到第一个就停止,避免重复添加
                            break
                ## 添加附属物
                label = "检修井"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelJXJ = pd.concat([excelJXJ, temp1], ignore_index=True)


                label = "消火栓"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelXHS = pd.concat([excelXHS, temp1], ignore_index=True)

                label = "检查井"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelJCJ = pd.concat([excelJCJ, temp1], ignore_index=True)
                label = "化粪池"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelHFC = pd.concat([excelHFC, temp1], ignore_index=True)

                label = "雨篦"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelYB = pd.concat([excelYB, temp1], ignore_index=True)

                label = "预留口"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelYLK = pd.concat([excelYLK, temp1], ignore_index=True)

                label = "地灯"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelDD = pd.concat([excelDD, temp1], ignore_index=True)

                label = "出地"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelCD = pd.concat([excelCD, temp1], ignore_index=True)

                label = "路灯"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelLD = pd.concat([excelLD, temp1], ignore_index=True)

                label = "配电箱"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelPDX = pd.concat([excelPDX, temp1], ignore_index=True)

                label = "接线箱"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelJXX = pd.concat([excelJXX, temp1], ignore_index=True)

                label = "控制柜"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelKZG = pd.concat([excelKZG, temp1], ignore_index=True)

                label = "信号灯"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelXHD = pd.concat([excelXHD, temp1], ignore_index=True)

                label = "手孔井"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelSKJ = pd.concat([excelSKJ, temp1], ignore_index=True)

                label = "电信发射塔"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelFST = pd.concat([excelFST, temp1], ignore_index=True)

                label = "人孔井"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelRKJ = pd.concat([excelRKJ, temp1], ignore_index=True)

                label = "摄像头"
                if label in dataAll[ii].loc[i, "附属物"]:
                    z = z1
                    temp1 = pd.DataFrame({"管线点预编号": dataAll[ii].loc[i, "管线点预编号"], "地面": dataAll[ii].loc[i, "地面"],
                                          "类型": dataAll[ii].loc[i, "附属物"], "x": dataAll[ii].loc[i, "X"],
                                          "y": dataAll[ii].loc[i, "Y"], "z": z},
                                         index=[0])
                    excelSXT = pd.concat([excelSXT, temp1], ignore_index=True)



        # 附属物去重
        excelJXJ = excelJXJ.drop_duplicates()
        excelXHS = excelXHS.drop_duplicates()
        excelJCJ = excelJCJ.drop_duplicates()
        excelHFC = excelHFC.drop_duplicates()
        excelYB = excelYB.drop_duplicates()
        excelYLK = excelYLK.drop_duplicates()
        excelDD = excelDD.drop_duplicates()
        excelCD = excelCD.drop_duplicates()
        excelLD = excelLD.drop_duplicates()
        excelPDX = excelPDX.drop_duplicates()
        excelJXX = excelJXX.drop_duplicates()
        excelKZG = excelKZG.drop_duplicates()
        excelXHD = excelXHD.drop_duplicates()
        excelSKJ = excelSKJ.drop_duplicates()
        excelFST = excelFST.drop_duplicates()
        excelRKJ = excelRKJ.drop_duplicates()
        excelSXT = excelSXT.drop_duplicates()

        ## 导出管线

        # 根据getData函数导出, 1 为给水, 2 为污水, 3 为雨水, 4为路灯, 5为电力, 6 为交通信号, 7为燃气, 8为电力通讯, 9 电信

        # 给水附属物只有检修井跟消火栓
        if ii == 1:
            with pd.ExcelWriter(output_file_name, mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
                excelJoutR.to_excel(writer, sheet_name="矩形给水", index=False)
                excelJoutC.to_excel(writer, sheet_name="圆形给水", index=False)
                excelJXJ.to_excel(writer, sheet_name="给水检修井", index=False)
                excelXHS.to_excel(writer, sheet_name="给水消火栓", index=False)
            print("----------------------1 给水完成------------------------")

        # 污水附属物只有检查井，化粪池
        if ii == 2:
            with pd.ExcelWriter(output_file_name, mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
                excelJoutR.to_excel(writer, sheet_name="矩形污水", index=False)
                excelJoutC.to_excel(writer, sheet_name="圆形污水", index=False)
                excelJCJ.to_excel(writer, sheet_name="污水检查井", index=False)
                excelHFC.to_excel(writer, sheet_name="污水化粪池", index=False)
            print("----------------------2 污水完成------------------------")

        # 雨水附属物： 检查井 雨篦 预留口 检修井
        if ii == 3:
            with pd.ExcelWriter(output_file_name, mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
                excelJoutR.to_excel(writer, sheet_name="矩形雨水", index=False)
                excelJoutC.to_excel(writer, sheet_name="圆形雨水", index=False)
                excelJXJ.to_excel(writer, sheet_name="雨水检修井", index=False)
                excelJCJ.to_excel(writer, sheet_name="雨水检查井", index=False)
                excelYB.to_excel(writer, sheet_name="雨水雨篦", index=False)
                excelYLK.to_excel(writer, sheet_name="雨水预留口", index=False)
            print("----------------------3 雨水完成------------------------")

        # 路灯附属物： 出地 地灯 检修井 路灯 配电箱
        if ii == 4:
            with pd.ExcelWriter(output_file_name, mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
                excelJoutR.to_excel(writer, sheet_name="矩形路灯", index=False)
                excelJoutC.to_excel(writer, sheet_name="圆形路灯", index=False)
                excelCD.to_excel(writer, sheet_name="路灯出地", index=False)
                excelDD.to_excel(writer, sheet_name="路灯地灯", index=False)
                excelJXJ.to_excel(writer, sheet_name="路灯检修井", index=False)
                excelLD.to_excel(writer, sheet_name="路灯路灯", index=False)
                excelPDX.to_excel(writer, sheet_name="路灯配电箱", index=False)
            print("----------------------4 路灯完成------------------------")

        # 电力附属物： 出地 检修井 接线箱 控制柜 配电箱
        if ii == 5:
            with pd.ExcelWriter(output_file_name, mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
                excelJoutR.to_excel(writer, sheet_name="矩形电力", index=False)
                excelJoutC.to_excel(writer, sheet_name="圆形电力", index=False)
                excelCD.to_excel(writer, sheet_name="电力出地", index=False)
                excelJXJ.to_excel(writer, sheet_name="电力检修井", index=False)
                excelJXX.to_excel(writer, sheet_name="电力接线箱", index=False)
                excelKZG.to_excel(writer, sheet_name="电力控制柜", index=False)
                excelPDX.to_excel(writer, sheet_name="电力配电箱", index=False)
            print("----------------------5 电力完成------------------------")

        # 交通信号附属物： 出地 检修井  控制柜 信号灯
        if ii == 6:
            with pd.ExcelWriter(output_file_name, mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
                excelJoutR.to_excel(writer, sheet_name="矩形交通信号", index=False)
                excelJoutC.to_excel(writer, sheet_name="圆形交通信号", index=False)
                excelCD.to_excel(writer, sheet_name="交通信号出地", index=False)
                excelJXJ.to_excel(writer, sheet_name="交通信号检修井", index=False)
                excelKZG.to_excel(writer, sheet_name="交通信号控制柜", index=False)
                excelXHD.to_excel(writer, sheet_name="交通信号信号灯", index=False)
            print("----------------------6 交通信号完成------------------------")

        # 燃气附属物： 检修井 预留口
        if ii == 7:
            with pd.ExcelWriter(output_file_name, mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
                excelJoutR.to_excel(writer, sheet_name="矩形燃气", index=False)
                excelJoutC.to_excel(writer, sheet_name="圆形燃气", index=False)
                excelJXJ.to_excel(writer, sheet_name="燃气检修井", index=False)
                excelYLK.to_excel(writer, sheet_name="燃气预留口", index=False)
            print("----------------------7 燃气完成------------------------")

        # 电力通讯附属物： 手孔井
        if ii == 8:
            with pd.ExcelWriter(output_file_name, mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
                excelJoutR.to_excel(writer, sheet_name="矩形电力通讯", index=False)
                excelJoutC.to_excel(writer, sheet_name="圆形电力通讯", index=False)
                excelSKJ.to_excel(writer, sheet_name="电力通讯手孔井", index=False)
            print("----------------------8 电力通讯完成------------------------")
        # 电信附属物： 电信发射塔 接线箱 人孔井 摄像头 手孔井 预留口
        if ii == 9:
            with pd.ExcelWriter(output_file_name, mode='a', engine="openpyxl", if_sheet_exists="replace") as writer:
                excelJoutR.to_excel(writer, sheet_name="矩形电信", index=False)
                excelJoutC.to_excel(writer, sheet_name="圆形电信", index=False)
                excelFST.to_excel(writer, sheet_name="电信发射塔", index=False)
                excelJXX.to_excel(writer, sheet_name="电信接线箱", index=False)
                excelRKJ.to_excel(writer, sheet_name="电信人孔井", index=False)
                excelSXT.to_excel(writer, sheet_name="电信摄像头", index=False)
                excelSKJ.to_excel(writer, sheet_name="电信手孔井", index=False)
                excelYLK.to_excel(writer, sheet_name="电信预留口", index=False)
            print("----------------------9 电信完成------------------------")

if __name__ == '__main__':
    path = r".\data\成果表预处理.xls"
    outpath = r".\data\out\深南大道.xlsx"
    inputdata = getData(path,outpath)

    # 在这里判断得出第7个表燃气的管径或断面类型为float, 会引起错误,所以要将所有管径或断面字段转为str类型
    # for i in range(1, len(inputdata)):
    #     print(type(inputdata[i]["管径或断面"][0]))
    # # print(type(getData(path)[0]))

    #测试偏移
    # print(inputdata[0]["X"],"-----", inputdata[0]["Y"])



    # outpath = r"H:\work\2022.8\model.xlsx"
    exportExcel(inputdata,outpath)