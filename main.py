# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import pandas as pd
import os
def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('PyCharm')
    a = "检查井"
    c = "检修井"
    b = "井"
    print(b in a)
    print(b in c)


    data = {

        'A': [1, 0, 1, 1],
        'B': [0, 2, 5, 0],
        'C': [4, 0, 4, 4],
        'D': [1, 0, 1, 1]
    }
    df = pd.DataFrame(data=data)
    # 默认保留第一次出现的重复项
    df1 = df.drop_duplicates()

    print(df)
    print(df1)
    import pandas as pd

    # df = pd.DataFrame([{'col1': 'a', 'col2': '1'}, {'col1': 'b', 'col2': '2'}])
    df = pd.DataFrame([{'col1': 'a', 'col2': 1.0}, {'col1': 'b', 'col2': 1.0}])

    print(type(df['col2'][0]))
    df['col2'] = df['col2'].astype('str')
    print(type(df['col2'][0]))
    #
    #
    #
    # df['col2'] = df['col2'].astype('float64')

    xxx = not ("X" in "100.0")
    print(xxx)

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
