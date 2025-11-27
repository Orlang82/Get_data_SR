import xlwings as xw
@xw.func
def py_RoundLR(data, threshold):
    """
    Возвращает 0 если отклонение не превышает заданный уровень Threshold.
    
    Параметры:
    data : число для проверки
    threshold : пороговое значение
    """
    if abs(data) > threshold:
        return data
    else:
        return 0