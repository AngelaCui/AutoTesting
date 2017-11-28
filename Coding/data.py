# -*- coding: utf-8 -*-
"""
Created on Tue Jun 13 14:17:51 2017

@author: admin
"""

# 通信协议解析
import math

# 数据区段
def dataprocess(maindata):

    while True:
        if maindata[0] != 0xa2:
            return 'Incorrect package'
        else:
            portnum = int(maindata[1]) *10 + int(maindata[2]) # 端口编号 # 问题：2 bytes 如何合并
            bytelen = maindata[3] # 实际长度字节数 # 问题：是否要与数据帧中数据长度想匹配
            if maindata[4] == 1:
                cons_stat = 1 # 1-处于施工状态
            elif maindata[4] == 0:
                cons_stat = 0 # 0-未施工
            else: 
                return 'Construction Status error'
            if maindata[5] >0 and maindata[5] <= 20:
                blocknum = maindata[5] # 闭塞分区数量
                # 5. 闭塞分区状态
                stat = {
                    0b0001:"正常占用",
                    0b0010:"空闲",
                    0b0011:"故障占用",  #协议中未011 而不是0011
                    0b0100:"失去分路",
                    0b0101:"出清（失去分路延时中）",
                    0b0110:"正常占用（越站调车）"}
                block_stat = [] # 依次为闭塞分区状态
                for i in range(int(blocknum/2)):
                    stat1 = stat[(maindata[6+i]>>4) & 0xF] # 字节内前一闭塞分区状态
                    stat2 = stat[maindata[6+i] & 0xF] # 字节内后一闭塞分区状态
                    block_stat.append(stat1)
                    block_stat.append(stat2)
                if blocknum % 2 == 1:
                    stat3 = stat[(maindata[6+int(blocknum/2)]>>4) & 0xF]
                    block_stat.append(stat3)
                # 闭塞分区状态结束时所处数据链中位置
                stat_end_num = 5 + math.ceil(blocknum/2)             
                
                # 6. 闭塞分区行车区间ID信息
                block_info = []
                for i in len(blocknum):
                    # 问题：无行车区间时，数据链中为0？所以无需分类讨论？
                    block_info.append(maindata[i + stat_end_num])
                # 闭塞分区行车区间信息结束时位置
                info_end_num = stat_end_num + blocknum
            # 若闭塞分区数量有误或为0    
            elif maindata[5] < 0 or maindata[5] > 20:
                return 'Exceeded block number'
            else:
                info_end_num = 5
                continue
            # 区间边界数量
            # 问题：数据包中此处如何显示？协议中边界数量为1-2，数据包中也是对应的吗？
            pos = info_end_num + 1
            edgenum = maindata[pos]
            
            # 边界1 
            edge_id = [] #行车区间ID
            edge_stat = []
            edge_SA = []
            edge_block = []
            edge_SA_stat = []
            
            e_stat = { # 边界行车区间状态
                0b001:"正常占用",
                0b010:"失去分路",
                0b011:"故障占用",
                0b100:"空闲"
                }
            e_SA = { # 边界信号许可类型
                0b00:"无信号许可",
                0b01:"发起信号许可",
                0b10:"应答信号许可",
                0b11:"故障（按无信号许可处理）"
                }
            e_block = { # 边界闭塞分区状态
                0b001:"正常占用",
                0b010:"失去分路",
                0b011:"故障占用",
                0b100:"空闲"
                }
            eSA_stat = { # 信号许可生成类型
                0b01:"新生成",
                0b10:"已确认"
                }
                
            for i in range(edgenum):
                edge_id.append(maindata[pos+1]) # 边界1 id
                edge_stat.append(e_stat[(maindata[pos+2])&0b111])
                edge_SA.append(e_SA[(maindata[pos+2]>>3)&0b11])
                edge_block.append(e_block[(maindata[pos+2]>>5)&0b111])
                edge_SA_stat.append(eSA_stat[(maindata[pos+3])&0b11])
                pos += 3
                
                
        return {"端口编号":portnum,"长度":bytelen,"施工状态":cons_stat,
                "闭塞分区数量":blocknum,"闭塞分区状态":block_stat,
                "闭塞分区行车区间信息":block_info,"边界行车区间信息":edge_id,
                "边界行车区间状态":[edge_stat,edge_SA,edge_block,edge_SA_stat]}