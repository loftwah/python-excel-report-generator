"""
Author: Roomey Rahman
mail: roomeyrahman@gmail.com"""


head = [
    {'title': 'A', 'key': 'a', 'style': {'font': {'font_size': '11', 'font_family': 'Calibri', 'bold': True, 'italic': False,
                      'underline': 'none', 'color': 'FF000000'},
             'alignment': {'horizontal': 'center', 'vertical': 'center'}}},
    {'title': 'B', 'key': 'b',
     'children':
         [
             {'title': 'C', 'key': 'c'},
             {'title': 'X', 'key': 'x'},
             {'title': 'D', 'key': 'd',
              'children': [
                  {'title': 'E', 'key': 'e'},
                  {'title': 'F', 'key': 'f'}
              ]
              }
         ],
     },
    {
        'title': 'G', 'key': 'g',
        'children': [
             {'title': 'H', 'key': 'h'},
             {'title': 'J', 'key': 'j'},
             {'title': 'I', 'key': 'i',
              'children': [
                  {'title': 'K', 'key': 'k'},
                  {'title': 'L', 'key': 'l',
                   'children': [
                        {'title': 'H', 'key': 'h'},
                        {'title': 'J', 'key': 'j'}]
                   }
              ]
              }
        ]
    },
    {
        'title': 'Z', 'key': 'z'
    }
]

dataframe = [
    {
        'a': 'Project Introduction',
        'b.c': 100,
        'b.x': 10,
        'b.d.e': 10,
        'b.d.f': 10,
        'g.h': 10,
        'g.j': 20,
        'g.i.k': 20,
        'g.i.l.h': 40,
        'g.i.l.j': 40,
        'z': 40
    },
    {
        'a': 'Project Introduction',
        'b.c': 100,
        'g.h': 10,
        'g.j': 20,
        'g.i.k': 20,
        'g.i.l.h': 40,
        'g.i.l.j': 40,
        'z': 40
    },
    {
        'a': 'Project Introduction',
        'b.c': 100,
        'b.x': 10,
        'b.d.e': 10,
        'b.d.f': 10,
        'g.i.k': 20,
        'g.i.l.h': 40,
        'g.i.l.j': 40,
        'z': 40
    }
]

import pandas as pd


class ExcelDataProcessing:
    def __init__(self, head, tableData, headType = 0, rowStart = 1):
        """
        :param head: will receive a list which will be the excel column headline data. Each item of the head will be a dictionary type.
        :param tableData: tabelData is the cell row value of excel report
        :param headType: headType is either 0 or any other number. if column head have no column information and user explicitly identify the column cell information then
        headType will be any other number except 0.
        """
        if headType == 0:
            self.headDepth = self.max_depth(head)
            self.head = self.headMap(head)
            (self.header, self.dataKeyMap) = self.headerPreparation(self.head, rowStart=rowStart)

        else:
            self.header = head

        if type(tableData) == dict:
            self.dataframe = pd.DataFrame(tableData)
        else:
            self.dataframe = self.dataMaping(tableData, self.dataKeyMap)


    def list_depth(self, List):
        """calculate the depth of the list item"""
        str_list = str(List)
        counter = 0
        for i in str_list:
            if i == "[":
                counter += 1
        return (counter)

    def max_depth(self, head):
        """calculate the max_depth of a nested list"""
        m_depth = 0
        for i in head:
            c = self.list_depth(i)
            m_depth = max(m_depth, c)
        return m_depth


    def childCount(self, headItem):
        """count the number of child of the tree structure"""
        sum = 0
        if 'children' in headItem:
            for i in headItem['children']:
                if 'children' not in i:
                    sum += 1
                else:
                    sum += self.childCount(i)
            return sum
        else:
            return 0


    def rowColSpan(self, item, index, maxRowSpan = 1, rowDepth = 0):
        """calculate rowspan, colspan, index and rowlevel of the tree"""
        if 'children' in item:
            maxRowSpan -= 1
            item['rowlevel'] = rowDepth
            rowDepth += 1
            for cIndex, i in enumerate(item['children']):
                self.rowColSpan(i, cIndex+index, maxRowSpan, rowDepth)
            item['colspan'] = self.childCount(item)
            item['rowspan'] = 1
            item['index'] = str(index) + ":" + str(index+(item['colspan']))
        else:
            item['rowspan'] = maxRowSpan
            item['colspan'] = 1
            item['rowlevel'] = rowDepth
            item['index'] = index
            index = index + item['colspan']


    def headMap(self, head):
        """Map head with metadata. Add aditional information(rowspan, colspan, index, rowlevel)"""
        # maximumDepth = self.headDepth
        for index, item in enumerate(head):
            if index == 0:
                self.rowColSpan(item, index=index, maxRowSpan=self.headDepth + 1)
            else:
                if type(head[index-1]['index']) == int:
                    newIndex = head[index-1]['index'] + 1
                else:
                    newIndex = int(head[index-1]['index'].split(':')[1])
                self.rowColSpan(item, index=newIndex, maxRowSpan=self.headDepth + 1)
        return head


    def cell_name(self, n):
        """find the cell name for excel"""
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string


    def createCell(self, item, rowStart=1):
        if type(item['index'])==int:
            item['column'] = self.cell_name(item['index']+1) + str(rowStart + item['rowlevel']) + ":" + self.cell_name(item['index']+item['colspan']) + str(rowStart + item['rowlevel'] + item['rowspan']-1)
        else:
            index = item['index'].split(':')
            item['column'] = self.cell_name(int(index[0]) + 1) + str(rowStart + item['rowlevel']) + ":" + self.cell_name(int(index[1])) + str(rowStart +  item['rowlevel'] + item['rowspan']-1)
        return item


    def headerPreparation(self, head, header = list(), rowStart = 1, dataKeyMap = dict(), parent = ''):
        for item in head:
            if 'style' in item:
                item['font'] = item['style'].get('font')
                item['alignment'] = item['style'].get('alignment')
            header.append(self.createCell(item, rowStart))

            if 'children' in item:
                self.headerPreparation(item['children'], header, rowStart, dataKeyMap, parent=(parent + item['key']+'.'))
            else:
                dataKeyMap[(parent + item['key'])] = list()

        return (header, dataKeyMap)


    def dataMaping(self, tableData, dataKeyMap):
        headKeys = list(dataKeyMap.keys())
        for item in tableData:
            for i in headKeys:
                if i in item:
                    self.dataKeyMap[i].append(item[i])
                else:
                    self.dataKeyMap[i].append('')

        dataframe = pd.DataFrame(self.dataKeyMap)
        return dataframe

