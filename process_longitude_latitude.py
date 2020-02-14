import numpy as np
import folium
import os
import pandas as pd


def main():
    data_url = "jingweidu2.xlsx"
    data = pd.read_excel(data_url, sheet_name='Sheet1')
    data_list = data.to_numpy()
    oil_and_speed_data_list = []
    for i in range(len(data_list)):
        l = []
        l.append(data_list[i][6])
        l.append(data_list[i][5])
        oil_and_speed_data_list.append(l)
    # print(oil_and_speed_data_list)
    m = folium.Map([25.99242, 119.367781], zoom_start=10)  # 中心区域的确定
    route = folium.PolyLine(  # polyline方法为将坐标用线段形式连接起来
        oil_and_speed_data_list,  # 将坐标点连接起来
        weight=3,  # 线的大小为3
        color='orange',  # 线的颜色为橙色
        opacity=0.8  # 线的透明度
    ).add_to(m)  # 将这条线添加到刚才的区域m内
    m.add_child(
        folium.ClickForMarker(popup='Waypoint')
    )
    m.save(os.path.join(r'C:\Users\87660\Desktop', 'chuli2.html'))  # 将结果以HTML形式保存到桌面上


if __name__ == '__main__':
    main()
