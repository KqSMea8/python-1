from init_fbi import *
import scipy as sp
import seaborn as sns
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['SimHei'] # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False # 用来正常显示负号
# run following command in jupyter to show plot inline
# %matplotlib inline
# run following command in jupyter to set retina display
# %config InlineBackend.figure_format = 'retina'
import functools
import random
import codecs
import sys
import re
import itertools
import enum

import pyecharts
import networkx
import igraph

import splinter
import requests
from bs4 import BeautifulSoup
# wxpython------------------------------------------------------------------------
def initWxPython():
    import wx

def fileDialog(message='Select File',
    wildcard='Excel 2007 (*.xlsx)|*.xlsx|Excel 2003 (*.xls)|*.xls',
    single=True):
    if 'app' in locals():
        del app
    app = wx.App()
    file_dialog = wx.FileDialog(parent=None, message=message, wildcard=wildcard,
        style=wx.FD_OPEN if single else wx.FD_MULTIPLE)
    file_dialog.ShowModal()
    file_name = file_dialog.GetPath()
    app.ExitMainLoop()
    return file_name

def excelSheetDialog(excelFile, title='Select Sheet(s)',
    single=True): # multiselect not working correctly
    def list_select(evt):
        global sheet_name
        selected_index = evt.GetEventObject().GetSelection()
        sheet_name = shts[selected_index]
        print(sheet_name)
        frame.Hide()
        app.ExitMainLoop()
    shts = getExcelSheets(excelFile)
    if 'app' in locals():
        del app
    app = wx.App()
    frame = wx.Frame(parent=None, title=title)
    frame.Center()
    panel = wx.Panel(parent=frame)
    listbox = wx.ListBox(parent=panel, choices=shts,
        style=wx.LB_SINGLE if single else wx.LB_MULTIPLE)
#     panel.Bind(wx.EVT_LISTBOX, list_select)
    panel.Bind(wx.EVT_LISTBOX_DCLICK, list_select)
    box = wx.BoxSizer()
    box.Add(listbox, proportion=1, flag=wx.EXPAND)
    panel.SetSizer(box)
    frame.Show()
    app.SetExitOnFrameDelete(True)
    app.MainLoop()
    return sheet_name

# Statistics------------------------------------------------------------------------
def factorial(n):
    '''
    return n!
    '''
    # method 1
    return functools.reduce(lambda x, y: x * y, range(1, n+1))
    # method 2
#     if n==1:
#         return n
#     else:
#         return n * factorial(n - 1)

def combination(n, m):
    '''
    return C_n^m = n! / m! / (n - m)!
    '''
    return factorial(n) / factorial(m) / factorial(n - m)

def mul(*args):
    result = args[0]
    for arg in args[1:]:
        result *= arg
    return result

def ecdf(data):
    '''
    计算数据的ECDF值
    ECDF (Empirical Cumulative Distribution Function)
        将数据从小到大排列，并用排名除以总数计算每个数据点在所有数据中的位置占比
        比如总共100个数据中排第20位的数据，其位置占比为20/100=0.2
    '''
    x = np.sort(data)
    y = np.arange(1, len(x)+1) / len(x)
    return (x, y)

def plot_ecdf(data, xlabel=None , ylabel='ECDF', label=None):
    '''
    绘制ECDF图
    '''
    x, y = ecdf(data)
    _ = plt.plot(x, y, marker='.', markersize=3, linestyle='none', label=label)
    _ = plt.legend(markerscale=4)
    _ = plt.xlabel(xlabel)
    _ = plt.ylabel(ylabel)
    plt.margins(0.02)

def cohen_d(data1, data2):
    '''
    Cohen's d，是均值的差值除以两个样本综合的标准差
        d = (x1.mean - x2.mean) / s_p
        s_p = sqrt(((n1 - 1) * s1^2 + (n2 - 1) * s2^2) / (n1 + n2 -2))
    Cohen's d的数值范围：当它的值为0.8代表有较大的差异，0.5位列中等，0.2较小，0.01则非常之小
    '''
    n1 = len(data1)
    n2 = len(data2)
    x1 = np.mean(data1)
    x2 = np.mean(data2)
    var1 = np.var(data1, ddof=1)
    var2 = np.var(data2, ddof=1)
    sp = np.sqrt(((n1 - 1) * var1 + (n2 - 1) * var2) \
        / (n1 + n2 - 2))
    return (x1-x2)/sp

def norm_pdf(x, mu, sigma):
    '''
    正态分布概率密度函数
    '''
    pdf = np.exp(-((x - mu)**2) / (2 * sigma**2)) / (sigma * np.sqrt(2 * np.pi))
    return pdf

def mean_ci(data, alpha=0.05):
    '''
    给定样本数据，计算均值1-alpha的置信区间
    '''
    sample_size = len(data)
    std = np.std(data, ddof=1)  # 估算总体的标准差
    se = std / np.sqrt(sample_size)  # 计算标准误差
    point_estimate = np.mean(data)
    z_score = sp.stats.norm.isf(alpha / 2)  # 置信度1-alpha
    confidence_interval = (point_estimate - z_score * se, point_estimate + z_score * se)
    return confidence_interval

# 自定义类------------------------------------------------------------------------
class TreeNode(object):
    '''
    TreeNode(name, value, children)
    children is a list of nodes
    '''
    def __init__(self, name, value=1, parent=None):
        super(TreeNode, self).__init__()
        self.name = name
        self.value = value
        self.parent = parent
        self.child = {}
    def __repr__(self):
        return 'TreeNode(%s)' % self.name
    def get_child(self, name, defval=None):
        return self.child.get(name, defval)
    def add_child(self, name, obj=None):
        if obj and not isinstance(obj, TreeNode):
            raise ValueError('TreeNode can only add another TreeNode obj as child')
        if obj is None:
            obj = TreeNode(name)
        obj.parent = self
        self.child[name] = obj
        return obj
    def del_child(self, name):
        if name in self.child:
            del self.child[name]
    def find_child(self, path, create=False):
        '''
        find child node by path / name, return None if not found
        path is a list, or a string like "name1->name2->name3"
        '''
        # convert path to a list if input is a string
        path = path if isinstance(path, list) else path.split('->')
        cur = self
        for sub in path:
            obj = cur.get_child(sub)
            if obj is None and create:
                obj = cur.add_child(sub)
            if obj is None:
                break
            cur = obj
        return obj
    def __contains__(self, item):
        return item in self.child
    def __len__(self):
        return len(self.child)
    def __bool__(self, item):
        return True
    def items(self):
        return self.child.items()
    def dump(self, indent=0):
        '''
        dump tree to string
        '''
        tab = '    ' * (indent - 1) + ' !- ' if indent > 0 else ''
        print('%s%s' % (tab, self.name))
        for name, obj in self.items():
            obj.dump(indent + 1)
    @property
    def path(self):
        '''
        return path string (from root to current node)
        '''
        if self.parent:
            return '%s->%s' % (self.parent.path.strip(), self.name)
        else:
            return self.name

class TreeNodeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, TreeNode):
            return {
                'name': obj.name,
                'value': obj.value,
                'children': [val for val in obj.child.values()],
            }
        return json.JSONEncoder.default(self, obj)

def TreeNodeDecoder(obj):
    if isinstance(obj, dict) and 'name' in obj:
        newobj = TreeNode(obj['name'])
        for name, subobj in obj.get('child', {}).items():
            newobj.add_child(name, TreeNodeDecoder(subobj))
        return newobj
    return obj

def TreeNodeToJson(tree):
    s = json.dumps(tree, cls=TreeNodeEncoder, indent=4,
        sort_keys=False, ensure_ascii=False)
    return json.loads(s)


def getTree(parent, df):
    '''
    get TreeNode iteratively from a dataframe
    example: getTree(TreeNode('root'), df)
    '''
    if df.shape[1]==1:
        for cell in df.iloc[:, 0]:
            parent.add_child(cell)
    else:
        for name, children in df.groupby(df.columns[0]):
            children = children.drop(df.columns[0], axis=1)
            child = parent.add_child(name)
            getTree(child, children)
    return parent

def getGraph(source, target, weight=1, category=None, directed=True):
    '''
    source, target, weight, category should be
    equal length lists or pandas Series
    weight and category (if not None) will be parsed as data
    '''
    # 准备edgelist
    edgelist = []
    source = list(source) if isinstance(source, pd.Series) else source
    target = list(target) if isinstance(target, pd.Series) else target
    weight = list(weight) if isinstance(weight, pd.Series) else weight
    category = list(category) if isinstance(category, pd.Series) else category
    for i in range(min(len(source), len(target))):
        edge = '|'.join([str(source[i]), str(target[i])])
        edge = '|'.join([edge, '1' if weight==1 else str(weight[i])])
        edge = '|'.join([edge, 'None' if category is None else str(category[i])])
        edgelist.append(edge)
    # 生成DiGraph
    dg = networkx.parse_edgelist(edgelist, delimiter='|',
        nodetype=str,
        create_using=networkx.DiGraph(),
        data=(('weight', float), ('category', str)))
    print('nodes:', len(dg.nodes()))
    print('edges:', len(dg.edges()))
    print('sources:', len(set(source)))
    print('targets:', len(set(target)))
    print('sample node:', list(dg.nodes())[0])
    print('sample edge:', list(dg.edges(data=True))[0])
    if directed:
        return dg
    else:
        return dg.to_undirected()

def getConnectedComponents(graph):
    '''
    return a dictionary:
        {'cc_nodes': count of nodes in each component,
        'components': list of connected components}
    '''
    if graph.is_directed():
        graph = graph.to_undirected()
    cc = list(networkx.connected_components(graph))
    print('connected components found:', len(cc))
    if len(cc)>0:
        cc_nodes = [len(comp) for comp in cc]
        print('max connected component size:', max(cc_nodes))
        return {'cc_nodes': cc_nodes,
            'components': cc}
    else:
        return None

def getKCommunityComponents(graph, k, algorithm=networkx.community.asyn_fluidc):
    '''
    return a dictionary:
        {'kc_nodes': count of nodes in each component,
        'components': list of k-community components}
    default algorithm: Asynchronous Fluid Communities algorithm
    '''
    if graph.is_directed():
        graph = graph.to_undirected()
    print('k communities to be found:', k)
    kc = list(algorithm(graph, k, max_iter=100))
    print('k communities found:', len(kc))
    if len(kc)>0:
        kc_nodes = [len(comp) for comp in kc]
        print('max k_community size:', max(kc_nodes))
        return {'kc_nodes': kc_nodes,
            'components': kc}
    else:
        return None

def getSimpleCycleComponents(graph):
    '''
    return a dictionary:
        {'sc_nodes': count of nodes in each component,
        'components': list of simple cycle components}
    '''
    sc = list(networkx.simple_cycles(graph))
    print('simple cycles found:', len(sc))
    if len(sc)>0:
        sc_nodes = [len(comp) for comp in sc]
        print('max simple cycle size:', max(sc_nodes))
        return {'sc_nodes': sc_nodes,
            'components': sc}
    else:
        return None

def getNodes(graph):
    '''
    return a list of nodes for pyecharts
    '''
    nodes = []
    for node in graph.nodes:
        nodes.append({'name': node})
    return nodes

def getLinks(graph):
    '''
    return a list of links for pyecharts
    '''
    links = []
    for link in graph.edges(data=True):
        links.append({'source': link[0],
            'target': link[1],
            'value': link[2]['weight']})
    return links
