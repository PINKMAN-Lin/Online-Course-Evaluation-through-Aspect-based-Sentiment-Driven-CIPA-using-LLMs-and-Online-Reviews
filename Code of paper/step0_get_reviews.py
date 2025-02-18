from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
import xlwt

if __name__ == '__main__':
    book = xlwt.Workbook(encoding='utf-8')
    sheet = book.add_sheet('sheet1', cell_overwrite_ok=True)
    sheet.write(0, 0, '用户昵称')
    sheet.write(0, 1, '评论内容')
    sheet.write(0, 2, '评论时间')
    sheet.write(0, 3, '点赞数')
    sheet.write(0, 4, '第几次课程')
    sheet.write(0, 5, '评分')
    row = 1
    driver = webdriver.Edge()  # 设置chrome驱动
    url = 'https://www.icourse163.org/course/BIT-268001'  # 设置要爬取的课程链接

    driver.get(url)
    ele = driver.find_element(value="review-tag-button", by=By.ID)
    # ele = driver.find_element_by_id("review-tag-button")  # 点击 课程评价
    ele.click()



    # xyy = driver.find_element_by_class_name("ux-pager_btn__next")  # 翻页功能
    for i in range(2025):  # 总页数

        connt = driver.page_source
        soup = BeautifulSoup(connt, 'html.parser')
        content = soup.find_all('div', {
            'class': 'ux-mooc-comment-course-comment_comment-list_item_body'})  # 全部评论

        for ctt in content:
            # 获取用户名
            user_name = ctt.find("a", {
                "class": "primary-link ux-mooc-comment-course-comment_comment-list_item_body_user-info_name"})
            user_name = user_name.text
            print(user_name)

            # 发布时间
            publish_time = ctt.find('div', {
                'class': 'ux-mooc-comment-course-comment_comment-list_item_body_comment-info_time'})
            publish_time = publish_time.text
            publish_time = publish_time[4:]
            print(publish_time)

            # 第几次课程
            course_nums = ctt.find('div', {
                'class': 'ux-mooc-comment-course-comment_comment-list_item_body_comment-info_term-sign'})
            course_nums = course_nums.text
            course_nums = course_nums.replace(" ", "")
            course_nums = course_nums.replace("n", "")
            print(course_nums)

            scontent = []
            aspan = ctt.find_all('span')
            for span in aspan:
                scontent.append(span.string)

            # 点赞数
            like = scontent[5]

            # 课程内容
            scontent = scontent[1]
            print(scontent)
            course_ratings = len(ctt.find_all('i', {
                'class': 'star ux-icon-custom-rating-favorite'}))
            sheet.write(row, 0, user_name)
            sheet.write(row, 1, scontent)
            sheet.write(row, 2, publish_time)
            sheet.write(row, 3, like)
            sheet.write(row, 4, course_nums)
            sheet.write(row, 5, course_ratings)
            row += 1
        xyy = driver.find_element(by=By.CLASS_NAME, value="ux-pager_btn__next")  # 翻页功能
        xyy.click()
    # 保存到Excel

    book.save('../2cost importance performanc analysis/data/python.xls')