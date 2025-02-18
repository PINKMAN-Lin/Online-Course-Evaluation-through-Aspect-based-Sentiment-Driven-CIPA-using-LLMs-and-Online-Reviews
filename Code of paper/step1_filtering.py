

import xlrd
import os
from langchain_openai import ChatOpenAI

from langchain.prompts import PromptTemplate
from langchain.chains import LLMChain
import asyncio

api_key = os.getenv("ZHIPU_API_KEY")
llm = ChatOpenAI(
        model="glm-4-flash",
        openai_api_key=api_key,  # 请填写您自己的APIKey
        openai_api_base="https://open.bigmodel.cn/api/paas/v4/"
    )


import xlwt

dataname = '人工智能4'
wb = xlrd.open_workbook("../data/mooc评论_" + dataname + ".xls")
sheet = wb.sheet_by_index(0)  # 通过索引的方式获取到某一个sheet，现在是获取的第一个sheet页，也可以通过sheet的名称进行获取，sheet_by_name('sheet名称')
rows = sheet.nrows  # 获取sheet页的行数，一共有几行
names = []
reviews = []
times = []
likes = []
class_order = []
rankings = []
for i in range(1, rows):
    names.append(str(sheet.cell(i, 0).value).strip())
    reviews.append(str(sheet.cell(i, 1).value).strip())
    times.append(str(sheet.cell(i, 2).value).strip())
    likes.append(str(sheet.cell(i, 3).value).strip())
    class_order.append(str(sheet.cell(i, 4).value).strip())
    rankings.append(str(sheet.cell(i, 5).value).strip())




template = '''
    帮我识别课程评论"{review}"是否有丰富的内容。
    这里的有丰富的内容，是相对于无丰富的内容而言的。
    有丰富的内容的清醒包括：评论中有关于课程的某个方面发表了某种情感倾向
    无丰富的内容的情形包括：
        1.空评论
        2.评论中没有中文字符
        3.评论虽然有中文字符，但缺少方面或者缺少情感倾向或者两者都缺。如“还可以 不错”，虽然具有情感倾向，但是不知道还可以不错究竟是对什么对象表达的还可以不错的观点，就不属于有丰富内容的评论。

    如果评论是无丰富的内容，请回复-1。
    如果评论是有丰富的内容，请回复1。
    
    '''
# template1_list = [template1%i for i in aspects]
prompt = PromptTemplate.from_template(template)
# print(prompt1)

chain = LLMChain(llm=llm, prompt=prompt)




book2 = xlwt.Workbook(encoding='utf-8')
sheet2 = book2.add_sheet('test', cell_overwrite_ok=True)
sheet2.write(0, 0, '用户昵称')
sheet2.write(0, 1, '评论内容')
sheet2.write(0, 2, '评论时间')
sheet2.write(0, 3, '点赞数')
sheet2.write(0, 4, '第几次课程')
sheet2.write(0, 5, '评分')


async def asy_chain_tool(chain, name, review, time, like, class_order, ranking):
  input = {'review': review}
  resp = await chain.ainvoke(input)
  # result = {review:resp}
  #
  # print(result)
  resp['name'] = name
  resp['time'] = time
  resp['like'] = like
  resp['class_order'] = class_order
  resp['ranking'] = ranking
  print(resp)
  return resp


async def asy_chain():
  tasks = [asy_chain_tool(chain, names[i], reviews[i], times[i], likes[i], class_order[i],rankings[i]) for i in range(len(reviews))]
  results_all = await asyncio.gather(*tasks)
  return results_all


results_all = asyncio.run(asy_chain())
print(results_all[:10])

row = 1
for i in range(len(results_all)):
    if results_all[i]['text'] != '-1':
        sheet2.write(row, 0, results_all[i]['name'])
        sheet2.write(row, 1, results_all[i]['review'])
        sheet2.write(row, 2, results_all[i]['time'])
        sheet2.write(row, 3, results_all[i]['like'])
        sheet2.write(row, 4, results_all[i]['class_order'])
        sheet2.write(row, 5, results_all[i]['ranking'])
        row += 1


book2.save("../data_after_preprocessing/" + dataname + ".xls")



#


