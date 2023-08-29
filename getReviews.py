import datetime
from github import Github  # pip install PyGithub

import xlwt  # pip install xlwt


def getReviews():
    # 表格初始化
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('sermant-pr-commits', cell_overwrite_ok=True)

    # github的token，可自行搜索教程获取账号token
    g = Github("token",
               verify=False)

    # 仓库地址
    repo = g.get_repo("huaweicloud/Sermant")
    # 获取所有的review，设定时间起始值
    reviews = repo.get_pulls_review_comments(since=datetime.datetime(2023, 8, 1))
    # 获取所有的pr
    prs = repo.get_pulls(state='all')

    index = 0  # 表格写入行
    for review in reviews:
        pr_user = ""  # pr的作者
        for pr in prs:
            if pr.url == review.pull_request_url:
                pr_user = pr.user.login
                print("this pr url：", pr.url)
                print("this pr user：", pr_user)
                break
        if review.user.login == pr_user:  # 排除pr作者的comment
            continue

        # 表格写入，第一列comment提出者，第二列comment内容，第三列comment的url（可选增加：pr的url，pr的作者）
        sheet.write(index, 0, review.user.login)
        sheet.write(index, 1, review.body)
        sheet.write(index, 2, review.url)
        # sheet.write(index, 3, review.pull_request_url)
        # sheet.write(index, 4, pr_user)

        index += 1
        print(review.user.login, review.body, review.url)

    # 保存表格，名字可更改
    savepath = "sermant-pr-commits-2023-xx.xls"
    book.save(savepath)


if __name__ == "__main__":
    getReviews()
