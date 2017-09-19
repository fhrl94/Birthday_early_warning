import datetime
'''
返回现在时间和预定时间相差的秒数，参数使用datetime.time来构造
如果今天超过预定时间，则返回距离明天预定时间相差的秒数
例如：
print(timer(datetime.time(8,00))) 误差应该是小于1s
'''
def timer(times):
    nowtime = datetime.datetime.now().time()
    if times > nowtime:
        return (times.hour - nowtime.hour) * 60 * 60 + (times.minute - nowtime.minute) * 60 + (
            times.second - nowtime.second)
    else:
        return (times.hour - nowtime.hour) * 60 * 60 + (times.minute - nowtime.minute) * 60 + (
            times.second - nowtime.second)+60*60*24
