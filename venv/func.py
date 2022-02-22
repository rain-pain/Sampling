import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import os.path
import time

def decorate(func):
    def wrapper(*args, **kwargs):
        start_time = time.time()
        func(*args, **kwargs)
        print("--- %s seconds ---" % (time.time() - start_time))
    return wrapper

def file_data():
    file_name = input('Введите имя файла: ')
    try:
        f = open(file_name, 'r')
        data = f.read().replace('\n', ' ').split(' ')
        f.close()
        return [float(elem) for elem in data], file_name
    except Exception as ex:
        print(ex, ' - введите корректное имя файла')
        return file_data()

def graph_save(x, y, x_graph, y_graph, file_name):
    plt.plot(x_graph, y_graph, color='m')
    plt.scatter(x, y, color='c')
    plt.xlabel('x')
    plt.ylabel('y')
    plt.title('Диаграмма рассеивания и линия регрессии')
    plt.savefig(f'./{file_name[:-4]}/Image_{file_name[:-4]}')
    plt.close()

def formulas(formula, result, name, file_name, formula2 =''):
    if not os.path.isdir(f'./{file_name[:-4]}/additional'):
        os.mkdir(f'./{file_name[:-4]}/additional')

    formula = formula[1:len(formula) - 1]
    fig = plt.figure()
    text = fig.text(0, -0.015, formula + f' = {result}' + formula2)
    dpi = 300
    fig.savefig(f'./{file_name[:-4]}/additional/{name}.png', dpi = dpi)

    bbox = text.get_window_extent()
    width, height = bbox.size / float(dpi) + 0.45
    width -= 0.3
    height -= 0.4
    fig.set_size_inches((width, height))
    dy = (bbox.ymin / float(dpi)) / height
    text.set_position((0, -dy))
    fig.savefig(f'./{file_name[:-4]}/additional/{name}.png', dpi = dpi)
    plt.close()
