
import xlrd
from xlutils.copy import copy

citys = []

class City:
    def __init__(self, name):
        self.county = []
        self.name = name

    #添加县
    def AddCounty(self, name):
        self.county.append(name)

    #查找地名，如果是本市的，返回[市，县]
    def FindPlace(self, address):
        for i in self.county:
            #已经找到县
            if address.find(i) >= 0:
                return [self.name, i]
        
        #没有找到，比较市是否可以找到
        if address.find(self.name) >= 0:
            return [self.name, self.name]
        
        return []

#行政区初始化
def InitCitys():
    #哈尔滨
    # a = City("哈尔滨")
    # a.AddCounty("道里区")
    # a.AddCounty("南岗区")
    # a.AddCounty("道外区")
    # a.AddCounty("平房区")
    # a.AddCounty("松北区")
    # a.AddCounty("香坊区")
    # a.AddCounty("呼兰区")
    # a.AddCounty("阿城区")
    # a.AddCounty("双城区")
    # a.AddCounty("依兰县")
    # a.AddCounty("方正县")
    # a.AddCounty("宾县")
    # a.AddCounty("巴彦县")
    # a.AddCounty("木兰县")
    # a.AddCounty("通河县")
    # a.AddCounty("延寿县")
    # a.AddCounty("尚志")
    # a.AddCounty("五常")
    # citys.append(a)

    #齐齐哈尔
    b = City("齐齐哈尔")
    # b.AddCounty("龙沙区")
    # b.AddCounty("建华区")
    # b.AddCounty("铁锋区")
    # b.AddCounty("昂昂溪区")
    # b.AddCounty("富拉尔基区")
    # b.AddCounty("碾子山区")
    # b.AddCounty("梅里斯达斡尔族区")
    b.AddCounty("龙江县")
    b.AddCounty("依安县")
    b.AddCounty("泰来县")
    b.AddCounty("甘南县")
    b.AddCounty("富裕县")
    b.AddCounty("克山县")
    b.AddCounty("克东县")
    b.AddCounty("拜泉县")
    b.AddCounty("讷河")
    citys.append(b)

    #鸡西 
    c = City("鸡西")
    # c.AddCounty("鸡冠区")
    # c.AddCounty("恒山区")
    # c.AddCounty("滴道区")
    # c.AddCounty("梨树区")
    # c.AddCounty("城子河区")
    # c.AddCounty("麻山区")
    c.AddCounty("鸡东县")
    c.AddCounty("虎林")
    c.AddCounty("密山")
    citys.append(c)

    #鹤岗
    d = City("鹤岗")
    # d.AddCounty("向阳区")
    # d.AddCounty("工农区")
    # d.AddCounty("南山区")
    # d.AddCounty("兴安区")
    # d.AddCounty("东山区")
    # d.AddCounty("兴山区")
    d.AddCounty("萝北县")
    d.AddCounty("绥滨县")
    citys.append(d)

    #双鸭山
    e = City("双鸭山")
    # e.AddCounty("尖山区")
    # e.AddCounty("岭东区")
    # e.AddCounty("四方台区")
    # e.AddCounty("宝山区")
    e.AddCounty("集贤县")
    e.AddCounty("友谊县")
    e.AddCounty("宝清县")
    e.AddCounty("饶河县")
    citys.append(e)

    #大庆
    f = City("大庆")
    # f.AddCounty("萨尔图区")
    # f.AddCounty("龙凤区")
    # f.AddCounty("让胡路区")
    # f.AddCounty("红岗区")
    # f.AddCounty("大同区")
    f.AddCounty("肇州县")
    f.AddCounty("肇源县")
    f.AddCounty("林甸县")
    f.AddCounty("杜尔伯特蒙古族自治县")
    # f.AddCounty("大庆高新技术产业开发区")
    citys.append(f)

    #伊春
    g = City("伊春")
    # g.AddCounty("伊春区")
    # g.AddCounty("南岔区")
    # g.AddCounty("友好区")
    # g.AddCounty("西林区")
    # g.AddCounty("翠峦区")
    # g.AddCounty("新青区")
    # g.AddCounty("美溪区")
    # g.AddCounty("金山屯区")
    # g.AddCounty("五营区")
    # g.AddCounty("乌马河区")
    # g.AddCounty("汤旺河区")
    # g.AddCounty("带岭区")
    # g.AddCounty("乌伊岭区")
    # g.AddCounty("红星区")
    # g.AddCounty("上甘岭区")
    g.AddCounty("嘉荫县")
    g.AddCounty("铁力")
    citys.append(g)

    #佳木斯
    h = City("佳木斯")
    # h.AddCounty("向阳区")
    # h.AddCounty("前进区")
    # h.AddCounty("东风区")
    # h.AddCounty("郊区")
    h.AddCounty("桦南县")
    h.AddCounty("桦川县")
    h.AddCounty("汤原县")
    h.AddCounty("同江")
    h.AddCounty("富锦")
    h.AddCounty("抚远")
    citys.append(h)

    #七台河 
    i = City("七台河")
    # i.AddCounty("新兴区")
    # i.AddCounty("桃山区")
    # i.AddCounty("茄子河区")
    i.AddCounty("勃利县")
    citys.append(i)

    #牡丹江
    j = City("牡丹江")
    # j.AddCounty("东安区")
    # j.AddCounty("阳明区")
    # j.AddCounty("爱民区")
    # j.AddCounty("西安区")
    j.AddCounty("林口县")
    # j.AddCounty("牡丹江经济技术开发区")
    j.AddCounty("绥芬河")
    j.AddCounty("海林")
    j.AddCounty("宁安")
    j.AddCounty("穆棱")
    j.AddCounty("东宁")
    citys.append(j)

    #黑河 
    k = City("黑河")
    # k.AddCounty("爱辉区")
    k.AddCounty("嫩江县")
    k.AddCounty("逊克县")
    k.AddCounty("孙吴县")
    k.AddCounty("北安")
    k.AddCounty("五大连池")
    citys.append(k)

    #绥化 
    l = City("绥化")
    # l.AddCounty("北林区")
    l.AddCounty("望奎县")
    l.AddCounty("兰西县")
    l.AddCounty("青冈县")
    l.AddCounty("庆安县")
    l.AddCounty("明水县")
    # l.AddCounty("绥棱县")
    l.AddCounty("安达")
    l.AddCounty("肇东")
    # l.AddCounty("海伦")
    citys.append(l)

    #大兴安岭
    m = City("大兴安岭")
    m.AddCounty("漠河")
    m.AddCounty("呼玛县")
    m.AddCounty("塔河县")
    # m.AddCounty("加格达奇区")
    # m.AddCounty("松岭区")
    # m.AddCounty("新林区")
    # m.AddCounty("呼中区")
    citys.append(m)

def FindPlace(address):
    #去掉地址中的#号
    index = address.find('#')
    address = address[:index]
    if index >= 0:
        address = address[:index]

    for i in citys:
        a = i.FindPlace(address)
        if(len(a) > 0):
            return a
    
    return []


if __name__  == '__main__':
    InitCitys()
    #打开excel
    hExcel = xlrd.open_workbook("2.xls")
    #选择sheet
    hSheet = hExcel.sheet_by_index(0)
    nRows = hSheet.nrows

    hExcelNew = copy(hExcel)
    hSheetNew = hExcelNew.get_sheet(0)

    for i in range(nRows):
        #查找第一行地址，如果找到，写入第二行
        a = FindPlace(hSheet.row_values(i)[0])
        if(len(a) > 0):
            hSheetNew.write(i, 1, a[0])
            hSheetNew.write(i, 2, a[1])

        hExcelNew.save('2.xls')
        

    
