# coding=utf-8
import os, sys, cx_Oracle, smtplib, datetime, csv, zipfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


def input_verify(fmdate, todate, output_type):  #用户输入验证模块
    db = cx_Oracle.connect('name/psw@*****')
    sqlfm = "select to_date('" + fmdate + "','yyyy-mm-dd hh24:mi:ss') from dual"
    cursor = db.cursor()
    try:
        cursor.execute(sqlfm)
        print  '开始时间验证OK'
    except Exception, e:
        print Exception, ":", e
        print '开始时间有误，请重新输入'
        return 'FALSE'

    sqlto = "select to_date('" + todate + "','yyyy-mm-dd hh24:mi:ss') from dual"
    cusror = db.cursor()
    try:
        cursor.execute(sqlto)
        print  '结束时间验证OK'
    except Exception, e:
        print Exception, ":", e
        print '结束时间有误，请重新输入'
        return 'FALSE'
    if output_type.lower() == 'csv' or output_type.lower() == 'txt':
        print '格式验证OK，取数中，请稍等......'
        return 'TRUE'
    else:
        print '格式输入有误，请重新输入（CSV/TXT不区分大小写）：'


def data_query(fmdate, todate, output_type): #数据查询模块
    db = cx_Oracle.connect('sss/sss_123456@10.202.4.97:1521/sss')
    sql = "select transit_zno, TO_CHAR(insert_tm,'YYYY-MM-DD HH24:MI:SS'), waybill_no, des_zno from tt_as_out_ident"
    where = " where TO_CHAR(insert_tm, 'yyyy-mm-dd hh24:mi:ss') between '" + fmdate + "' and '" + todate + "'"
    cursor = db.cursor()
    cursor.execute(sql + where)
    result = cursor.fetchall()
    if result and output_type.lower() == 'csv':
        trans_result(result, 'csv')
    if result and output_type.lower() == 'txt':
        trans_result(result, 'txt')
    else:
        print '条件输入查询数据为空，程序退出，谢谢'
        sys.exit()


def get_current_dir(output_type):  #获取当前新建取数文件目录
    datafile_path = str(os.getcwd())
    datafile_name = datetime.datetime.now().strftime("%Y%m%d%H%M%S") + '.' + output_type
    datafile_name_path = datafile_path +"\\"+ datafile_name
    return datafile_name_path


def create_txtdatafile(output_type):  #新建txt头文件
    datafile_dir = get_current_dir(output_type)
    txt_header = "场地代码-时间段-运单号-目的地\n"
    f = open(datafile_dir, 'wb')
    f.write(txt_header)
    f.close()
    return datafile_dir


def add_data(datafile_dir, data): #写入取数文件CSV数据
    f = open(datafile_dir, 'a')
    f.write(data)
    f.close()


def add_csvdata(datafile_dir, data):   #写入取数文件CSV数据
    csvfile = open(datafile_dir, 'ab+')
    writer = csv.writer(csvfile)
    writer.writerows(data)
    csvfile.close()


def mail_to(filename):  #邮件发送模块
    user = "mail"
    pwd = "psw"
    to = "mail"

    msg = MIMEMultipart()
    msg["Subject"] = "取数"
    msg["From"] = user
    msg["To"] = to

    part = MIMEText("取数结果详见附件")
    msg.attach(part)
    part = MIMEApplication(open(filename, "rb").read())
    part.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(part)

    s = smtplib.SMTP_SSL("smtp.qq.com", 465)
    s.login(user, pwd)
    s.sendmail(user, to, msg.as_string())
    print "取数邮件已经发送，请注意查收，谢谢"
    s.close()
    sys.exit()


def trans_result(result, output_format): #取数格式转换
    dir = str(os.getcwd())
    now_time = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    if output_format.lower() == 'txt':
        datafile_dir = create_txtdatafile('txt')
        for each_item in result:
            line_result = each_item[0] + '-' + each_item[1] + '-' + each_item[2] + '-' + each_item[3] + '\n'
            add_data(datafile_dir, line_result)
            zip_file_name = file_zip(datafile_dir)
            print zip_file_name
            mail_to(zip_file_name)
    else:
        datafile_dir = get_current_dir('csv')
        csv_header = "场地代码,时间段,运单号,目的地\n"
        add_data(datafile_dir, csv_header)
        add_csvdata(datafile_dir, result)
        zip_file_name = file_zip(datafile_dir)
        mail_to(zip_file_name)


def file_zip(filename): #文件压缩模块
    zip_file_name = filename + '.zip'
    # f = zipfile.ZipFile(zip_file_name,'w',zipfile.ZIP_DEFLATED)
    f = zipfile.ZipFile(zip_file_name, 'w')
    f.write(filename)
    f.close()
    return zip_file_name


if __name__ == "__main__": #输入查询条件
    while True:
        # r_fmdate=str(raw_input('请输入开始时间 例：2016-12-12 23:59:59：'))
        _fmdate = '2015-12-11 00:00:00'
        # r_todate=str(raw_input('请输入结束时间 例：2016-12-12 23:59:59：'))
        _todate = '2016-12-12 00:00:00'
        # r_outpt_type=str(raw_input('请输入导出格式 CSV/TXT:'))
        _output_type = 'txt'
        if input_verify(_fmdate, _todate, _output_type) == 'TRUE':
            data_query(_fmdate, _todate, _output_type)
        else:

            continue

