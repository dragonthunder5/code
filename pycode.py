import win32com.client 
import pandas as pd
import time
import datetime
from dateutil.relativedelta import relativedelta
import calendar
from IPython.display import display, HTML
import co_job
import datetime


# get today's date in the format 22.02.2023
today = datetime.datetime.today().strftime('%d.%m.%Y')

global x , y ,z


def temp():
    job_list = [
        {'job': 'ZSD0049_AB_BW_COPABOOK_OUTBOUND', 'status':'Finished', 'comment':  co_job.sm37('ZSD0049_AB_BW_COPABOOK_OUTBOUND')},
        {'job': 'SII_CO0962_AB_DWH_COPA_BOOK_9462', 'status':'' , 'comment': co_job.sm37('SII_CO0962_AB_DWH_COPA_BOOK_9462')},
        {'job': 'SBT_FI0005_AR_19_YYFIESP_EXTRA', 'status': '', 'comment': co_job.sm37('SBT_FI0005_AR_19_YYFIESP_EXTRA')},
    ]

    html = '<table style="border-collapse: collapse; border: 1px solid black;">\n'
    html += '<thead>\n'
    html += '<tr style="background-color: yellow;">\n'
    html += '<th style="border: 1px solid black; padding: 8px;">CO jobs</th>\n'
    html += '<th style="border: 1px solid black; padding: 8px;">Status</th>\n'
    html += '<th style="border: 1px solid black; padding: 8px;">Comment</th>\n'
    html += '</tr>\n'
    html += '</thead>\n'
    html += '<tbody>\n'

    for i, job in enumerate(job_list, 1):
        
        if job['comment'] == 'Finished':
            status_color = 'green'
        elif job['comment'] == 'Error':
            status_color = 'yellow'
        else:
            status_color = 'red'
        html += '<tr>\n'
        html += f'<td style="border: 1px solid black; padding: 8px;">{job["job"]}</td>\n'
        html += f'<td style="border: 1px solid black; padding: 8px;"><b><span style="font-size:20.0pt;font-family:Wingdings;mso-fareast-font-family:Calibri;mso-fareast-theme-font:minor-latin;color:{status_color}">l<o:p></o:p></span></b></td>\n'
  
        html += f'<td style="border: 1px solid black; padding: 8px;">{job["comment"]}</td>\n'
        html += '</tr>\n'

    html += '</tbody>\n'
    html += '</table>\n'

    # create outlook email
    outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = 'mail@me.com'
    mail.CC = 'mail@me.com'
    mail.Subject = 'Job Monitoring CO (E1P-100 / E1P-022)'
    #mail.Attachments.Add(Source=r"C:\UserData\z004mfjs\Documents\abc.xlsx")

    # modify the mail body as per need
    mail.BodyFormat = 2

    intro = "<p>Hello Team,<br> As per provided revised job list. We have monitored jobs.</p>\n"
    #bye = "<br><br><br>Thanks and regards,<br>Srikanth<br>
    mail.HTMLBody = html
    x=html
    return x 
    #ail.Display()
    
    

#part1 = temp()

def combiner():
    import pandas as pd

    # Read the Excel file into a DataFrame
    df = pd.read_excel('parent_job.xlsx', sheet_name='Sheet1', header=None)

    # Find the column index of the job name





    #s=read.sort_values('STATUS',ascending=(False))
    filename = (r"C:\UserData\z004mfjs\Documents\sm37\co\co_frequency_job_status.xlsx")
    excel=pd.read_excel(filename,index_col=0)
     #print(excel)
    list2 = excel.index.values.tolist()

    excel1=pd.read_excel(filename,index_col=1)
    list3= excel1.index.values.tolist()


    lsparent=[]
    lsjob=[]
    stat=[]

    for i in range(len(list2)):
        job_name= list2[i]
        status=list3[i]
        string_cols = df.select_dtypes(include='object').columns
        string_contains = df[string_cols].apply(lambda x: x.str.contains(job_name, na=False))
        col_index = string_contains.any().idxmax()

        # Get the header of the column index
        header = df.iloc[0, col_index] 
        print(job_name,header)
        lsparent.append(header)
        lsjob.append(job_name)
        stat.append(status)
        


    # print(df)
     #print(df.shape[0])
    df= df.dropna()
    df.to_clipboard(index=False, header=False)
    #print(df)
    df=pd.DataFrame()
    df['Parent']=lsparent
    df['Job name']=lsjob
    df['Status']=stat

    df.to_excel('combined.xlsx',index= False)
    
    
def mail100():
   
    df1 = pd.read_excel(r'C:\UserData\z004mfjs\Documents\sm37\co\combined.xlsx')
                    
    html_table = df1.to_html(index=False)
    #print(df1)
    par=df1["Parent"].values
    jobn=df1["Job name"].values
    status=df1["Status"].values
    
    data = {
        "parent":par,
      "JOB NAME": jobn,
      "STATUS": status
    }
    
    x = data["STATUS"]
    y= data["JOB NAME"]
    z=data["parent"]
    
    f=[]
    a=[]
    c=[]
    r=[]
    others=[]
    
    
    # For 5620-RA/Variance/Settlement Jobs
    a_5620_ra_vs_jobs = []
    c_5620_ra_vs_jobs = []
    f_5620_ra_vs_jobs = []
    others_5620_ra_vs_jobs = []
    
    # For 5620-POC Jobs
    a_5620_poc_jobs = []
    c_5620_poc_jobs = []
    f_5620_poc_jobs = []
    others_5620_poc_jobs = []
    
    # For 4433-RA/Variance/Settlement Jobs
    a_4433_ra_vs_jobs = []
    c_4433_ra_vs_jobs = []
    f_4433_ra_vs_jobs = []
    others_4433_ra_vs_jobs = []
    
    # For 9461-RA/Variance/Settlement Jobs
    a_9461_ra_vs_jobs = []
    c_9461_ra_vs_jobs = []
    f_9461_ra_vs_jobs = []
    others_9461_ra_vs_jobs = []
    
    # For 9461-POC Jobs
    a_9461_poc_jobs = []
    c_9461_poc_jobs = []
    f_9461_poc_jobs = []
    others_9461_poc_jobs = []
    
    # For 9462-RA/Variance/Settlement Jobs
    a_9462_ra_vs_jobs = []
    c_9462_ra_vs_jobs = []
    f_9462_ra_vs_jobs = []
    others_9462_ra_vs_jobs = []

 



    for i in range(len(x)):
        if z[i] == "5620-RA/Variance/Settlement Jobs":
            if x[i] == "Finished":
                f_5620_ra_vs_jobs.append(y[i])
            elif x[i] == "Active":
                a_5620_ra_vs_jobs.append(y[i])
            elif x[i] == "Canceled":
                c_5620_ra_vs_jobs.append(y[i])
            else:
                others_5620_ra_vs_jobs.append(y[i])
        elif z[i] == "5620-POC Jobs":
            if x[i] == "Finished":
                f_5620_poc_jobs.append(y[i])
            elif x[i] == "Active":
                a_5620_poc_jobs.append(y[i])
            elif x[i] == "Canceled":
                c_5620_poc_jobs.append(y[i])
            else:
                others_5620_poc_jobs.append(y[i])
        elif z[i] == "4433-RA/Variance/Settlement Jobs":
            if x[i] == "Finished":
                f_4433_ra_vs_jobs.append(y[i])
            elif x[i] == "Active":
                a_4433_ra_vs_jobs.append(y[i])
            elif x[i] == "Canceled":
                c_4433_ra_vs_jobs.append(y[i])
            else:
                others_4433_ra_vs_jobs.append(y[i])
        elif z[i] == "9461-RA/Variance/Settlement Jobs":
            if x[i] == "Finished":
                f_9461_ra_vs_jobs.append(y[i])
            elif x[i] == "Active":
                a_9461_ra_vs_jobs.append(y[i])
            elif x[i] == "Canceled":
                c_9461_ra_vs_jobs.append(y[i])
            else:
                others_9461_ra_vs_jobs.append(y[i])
        elif z[i] == "9461-POC Jobs":
            if x[i] == "Finished":
                f_9461_poc_jobs.append(y[i])
            elif x[i] == "Active":
                a_9461_poc_jobs.append(y[i])
            elif x[i] == "Canceled":
                c_9461_poc_jobs.append(y[i])
            else:
                others_9461_poc_jobs.append(y[i])
        elif z[i] == "9462-RA/Variance/Settlement Jobs":
            if x[i] == "Finished":
                f_9462_ra_vs_jobs.append(y[i])
            elif x[i] == "Active":
                a_9462_ra_vs_jobs.append(y[i])
            elif x[i] == "Canceled":
                c_9462_ra_vs_jobs.append(y[i])
            else:
                others_9462_ra_vs_jobs.append(y[i])

           
        
    #print(f)    
    #removing duplicates
    f = list(dict.fromkeys(f))
    a = list(dict.fromkeys(a))
    c = list(dict.fromkeys(c))
    r = list(dict.fromkeys(r))
    others = list(dict.fromkeys(others))
    
    

    
    
    #print(f)
    #making the list to print one below the other    
    pd.options.display.max_rows = 999999999
    testf = "<br>".join(map(str,f))
    testa = "<br>".join(map(str,a))
    testc = "<br>".join(map(str,c))
    testr = "<br>".join(map(str,r))
    testo = "<br>".join(map(str,others))
    
    
    

    import datetime
    
    today = datetime.date.today()
    first_day = datetime.date(today.year, today.month, 1)
    
    working_days = []
    for i in range(31):
        day = first_day + datetime.timedelta(days=i)
        if len(working_days) < 7 and day.weekday() < 5:
            working_days.append(day.day)
        elif len(working_days) == 7:
            break
    
    print("The first 7 working days of the month are:", working_days)
    
    #print(working_days[0])
    
    day=today.day
    
    print("today is : "+ str(day))


    
    f_5620_ra_vs_jobs = list(dict.fromkeys(f_5620_ra_vs_jobs))
    a_5620_ra_vs_jobs = list(dict.fromkeys(a_5620_ra_vs_jobs))
    c_5620_ra_vs_jobs = list(dict.fromkeys(c_5620_ra_vs_jobs))
    #r_5620_ra_vs_jobs = list(dict.fromkeys(r_5620_ra_vs_jobs))
    others_5620_ra_vs_jobs = list(dict.fromkeys(others_5620_ra_vs_jobs))
    
    
    pd.options.display.max_rows = 999999999
    f_5620_ra_vs_jobs = "<br>".join(map(str, f_5620_ra_vs_jobs))
    a_5620_ra_vs_jobs = "<br>".join(map(str, a_5620_ra_vs_jobs))
    c_5620_ra_vs_jobs = "<br>".join(map(str, c_5620_ra_vs_jobs))
    #r_5620_ra_vs_jobs = "<br>".join(map(str, r_5620_ra_vs_jobs))
    others_5620_ra_vs_jobs = "<br>".join(map(str, others_5620_ra_vs_jobs))
   
    
    if a_5620_ra_vs_jobs :
        bodya ="<html><body>" "<br>these are with status <b> Active</b>:<br>" +" <span style='color: blue'>%s</span>" "<html><body>"%a_5620_ra_vs_jobs
    else:
        bodya = ""
    if c_5620_ra_vs_jobs:
        bodyc = "<br>these are with status <b> Canceled</b>:<br>" +" <span style='color: blue'>%s</span>" % c_5620_ra_vs_jobs
    else:
        bodyc = ""
    if others_5620_ra_vs_jobs:
        if working_days[2] <= day <= working_days[3]:
            bodo = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
           bodyo= "<br>these are with status <b> Did not run</b>:<br>" + " <span style='color: blue'>%s</span>" % others_5620_ra_vs_jobs
   
    else:
       bodyo= ""    
    # create the DataFrame with the specified column headings and row values
    # create a list of row values for the DataFrame
   
    
   
    
    
    f_5620_poc_jobs = list(dict.fromkeys(f_5620_poc_jobs))
    a_5620_poc_jobs = list(dict.fromkeys(a_5620_poc_jobs))
    c_5620_poc_jobs = list(dict.fromkeys(c_5620_poc_jobs))
    #r_5620_poc_jobs = list(dict.fromkeys(r_5620_poc_jobs))
    others_5620_poc_jobs = list(dict.fromkeys(others_5620_poc_jobs))
    
    # print(f)
    # making the list to print one below the other
    pd.options.display.max_rows = 999999999
  
    
    f_5620_poc_jobs = "<br>".join(map(str, f_5620_poc_jobs))
    a_5620_poc_jobs = "<br>".join(map(str, a_5620_poc_jobs))
    c_5620_poc_jobs = "<br>".join(map(str, c_5620_poc_jobs))
    #r_5620_poc_jobs = "<br>".join(map(str, r_5620_poc_jobs))
    others_5620_poc_jobs = "<br>".join(map(str, others_5620_poc_jobs))
    

        
    
    
    
    if a_5620_poc_jobs:
        body1a = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % a_5620_poc_jobs
    else:
        body1a = ""
    if c_5620_poc_jobs:
        body1c = "<br>these are with status <b> Cancelled</b>:<br>" + " <span style='color: blue'>%s</span>" % c_5620_poc_jobs
    else:
        body1c = ""
    if others_5620_poc_jobs:
        if working_days[2] <= day <= working_days[3]:
            body1o = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
           body1o= "<br>these are with status <b> Did not run</b>:<br>" + " <span style='color: blue'>%s</span>" % others_5620_poc_jobs
   
    else:
       body1o= ""    
        
        
        
        
    f_4433_ra_vs_jobs = list(dict.fromkeys(f_4433_ra_vs_jobs))
    a_4433_ra_vs_jobs = list(dict.fromkeys(a_4433_ra_vs_jobs))
    c_4433_ra_vs_jobs = list(dict.fromkeys(c_4433_ra_vs_jobs))
    #r_4433_ra_vs_jobs = list(dict.fromkeys(r_4433_ra_vs_jobs))
    others_4433_ra_vs_jobs = list(dict.fromkeys(others_4433_ra_vs_jobs))
    #print(f)
    #making the list to print one below the other
    
    pd.options.display.max_rows = 999999999
    
    f_4433_ra_vs_jobs = "<br>".join(map(str, f_4433_ra_vs_jobs))
    a_4433_ra_vs_jobs = "<br>".join(map(str, a_4433_ra_vs_jobs))
    c_4433_ra_vs_jobs = "<br>".join(map(str, c_4433_ra_vs_jobs))
    #r_4433_ra_vs_jobs = "<br>".join(map(str, r_4433_ra_vs_jobs))
    others_4433_ra_vs_jobs = "<br>".join(map(str, others_4433_ra_vs_jobs))
    
    if a_4433_ra_vs_jobs:
        body2a = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % a_4433_ra_vs_jobs
    else:
        body2a = ""
    if c_4433_ra_vs_jobs:
        body2c = "<br>these are with status <b> Canceled</b>:<br>" + " <span style='color: blue'>%s</span>" % c_4433_ra_vs_jobs
    else:
        body2c = ""
        
    if others_4433_ra_vs_jobs:
        if working_days[2] <= day <= working_days[5]:
            body2o = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
           body2o= "<br>these are with status <b> Did not run</b>:<br>" + " <span style='color: blue'>%s</span>" % others_4433_ra_vs_jobs
   
    else:
       body2o= ""
    
   

        
        
    f_9461_ra_vs_jobs = list(dict.fromkeys(f_9461_ra_vs_jobs))
    a_9461_ra_vs_jobs = list(dict.fromkeys(a_9461_ra_vs_jobs))
    c_9461_ra_vs_jobs = list(dict.fromkeys(c_9461_ra_vs_jobs))
    # r_9461_ra_vs_jobs = list(dict.fromkeys(r_9461_ra_vs_jobs))
    others_9461_ra_vs_jobs = list(dict.fromkeys(others_9461_ra_vs_jobs))
    
    # making the list to print one below the other
    pd.options.display.max_rows = 999999999
    
    # converting lists to strings with line breaks
    f_9461_ra_vs_jobs = "<br>".join(map(str, f_9461_ra_vs_jobs))
    a_9461_ra_vs_jobs = "<br>".join(map(str, a_9461_ra_vs_jobs))
    c_9461_ra_vs_jobs = "<br>".join(map(str, c_9461_ra_vs_jobs))
    # r_9461_ra_vs_jobs = "<br>".join(map(str, r_9461_ra_vs_jobs))
    others_9461_ra_vs_jobs = "<br>".join(map(str, others_9461_ra_vs_jobs))
    
    # generating the HTML body content based on job status
    if a_9461_ra_vs_jobs:
        body3a = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % a_9461_ra_vs_jobs
    else:
        body3a = ""
    if c_9461_ra_vs_jobs:
        body3c = "<br>these are with status <b> DID NOT RUN</b>:<br>" + " <span style='color: blue'>%s</span>" % c_9461_ra_vs_jobs
    else:
        body3c = ""
    if others_9461_ra_vs_jobs:
        if working_days[2] <= day <= working_days[3]:
            body3o = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
           body3o= "<br>these are with status <b> Did not run</b>:<br>" + " <span style='color: blue'>%s</span>" % others_9461_ra_vs_jobs
   
    else:
       body3o= ""
        
    
    
    
    
    
    f_9461_poc_jobs = list(dict.fromkeys(f_9461_poc_jobs))
    a_9461_poc_jobs = list(dict.fromkeys(a_9461_poc_jobs))
    c_9461_poc_jobs = list(dict.fromkeys(c_9461_poc_jobs))
    #r_9461_poc_jobs = list(dict.fromkeys(r_9461_poc_jobs))
    others_9461_poc_jobs = list(dict.fromkeys(others_9461_poc_jobs))
    
    pd.options.display.max_rows = 999999999
    
    f_9461_poc_jobs = "<br>".join(map(str, f_9461_poc_jobs))
    a_9461_poc_jobs = "<br>".join(map(str, a_9461_poc_jobs))
    c_9461_poc_jobs = "<br>".join(map(str, c_9461_poc_jobs))
    #r_9461_poc_jobs = "<br>".join(map(str, r_9461_poc_jobs))
    others_9461_poc_jobs = "<br>".join(map(str, others_9461_poc_jobs))
    
    if a_9461_poc_jobs:
        body4a = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % a_9461_poc_jobs
    else:
        body4a = ""
    if c_9461_poc_jobs:
        body4c = "<br>these are with status <b> DID NOT RUN</b>:<br>" + " <span style='color: blue'>%s</span>" % c_9461_poc_jobs
    else:
        body4c = ""
    
    if others_9461_poc_jobs:
        if working_days[2] <= day <= working_days[3]:
            body4o = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
           body4o= "<br>these are with status <b> Did not run</b>:<br>" + " <span style='color: blue'>%s</span>" % others_9461_poc_jobs
   
    else:
       body4o= ""
    




    f_9462_ra_vs_jobs = list(dict.fromkeys(f_9462_ra_vs_jobs))
    a_9462_ra_vs_jobs = list(dict.fromkeys(a_9462_ra_vs_jobs))
    c_9462_ra_vs_jobs = list(dict.fromkeys(c_9462_ra_vs_jobs))
    # r_9462_ra_vs_jobs = list(dict.fromkeys(r_9462_ra_vs_jobs))
    others_9462_ra_vs_jobs = list(dict.fromkeys(others_9462_ra_vs_jobs))
    
    # print(f)
    # making the list to print one below the other
    pd.options.display.max_rows = 999999999
    
    
    f_9462_ra_vs_jobs = "<br>".join(map(str, f_9462_ra_vs_jobs))
    a_9462_ra_vs_jobs = "<br>".join(map(str, a_9462_ra_vs_jobs))
    c_9462_ra_vs_jobs = "<br>".join(map(str, c_9462_ra_vs_jobs))
    # r_9462_ra_vs_jobs = "<br>".join(map(str, r_9462_ra_vs_jobs))
    others_9462_ra_vs_jobs = "<br>".join(map(str, others_9462_ra_vs_jobs))
    
    if a_9462_ra_vs_jobs:
        body5a = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % a_9462_ra_vs_jobs
    else:
        body5a = ""
    if c_9462_ra_vs_jobs:
        body5c = "<br>these are with status <b>Canceled</b>:<br>" + " <span style='color: blue'>%s</span>" % c_9462_ra_vs_jobs
    else:
        body5c = ""
    if others_9462_ra_vs_jobs:
        if working_days[2] <= day <= working_days[3]:
            body5o = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
           body5o= "<br>these are with status <b> Did not run</b>:<br>" + " <span style='color: blue'>%s</span>" % others_9462_ra_vs_jobs
   
    else:
       body5o= ""
    
    
    #lean jobs
    #import pandas as pd

    # Load the data from export1 and export2 into separate dataframes
    export1 = pd.read_excel('export1.xlsx')
    export2 = pd.read_excel('export2.xlsx')

    # Concatenate the two dataframes into a single dataframe
    combined = pd.concat([export1, export2], ignore_index=True)

    # Group the data by company code and fiscal year, and sum the number of entries
    summed = combined.groupby(['Company Code', 'Fiscal Year'])['Number of Entries'].sum().reset_index()

    # Print the final dataframe
    #print(summed)



    # Filter the data to include only company code 5620
    code_5620 = summed[summed['Company Code'] == 5620]
    code_8620 = summed[summed['Company Code'] == 8620]
    code_9461 = summed[summed['Company Code'] == 9461]
    code_9462 = summed[summed['Company Code'] == 9462]

    lean= '''<!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <title>Outlook email</title>
      </head>
      <body>
        <p>Records with blank subsequent documents:This will be re-processed during daily run.<br>During weekly clean-up process:<br>
        1. Re-processed records with blank subsequent document
        <br>2.Report errors &amp; Resolution: Example errors - PC locked, Object Locked, derivation rule error etc.
        </p>
      </body>
    </html>
    '''
    # Print the sum of entries for company code 5620
    #print("The number of entries for company code 5620 is:", code_5620['Number of Entries'].sum())

    #print(code_5620['Number of Entries'].sum())
    
    
 #create HTML table
    job_list = [
        {'job': '5620-RA/Variance/Settlement Jobs', 'status': 'Finished', 'comment': bodya+bodyc+bodyo},
        {'job': '5620-POC Jobs', 'status': 'Finished', 'comment': body1a+body1c+body1o},
        {'job': '4433-RA/Variance/Settlement Jobs', 'status': 'Finished', 'comment':body2a+body2c+body2o},
        {'job': '9461-RA/Variance/Settlement Jobs', 'status': 'Finished', 'comment': body3a+body3c+body3o},
        {'job': '9461-POC Jobs', 'status': 'Finished', 'comment': body4a+body4c+body4o},
        {'job': '9462-RA/Variance/Settlement Jobs', 'status': 'Finished', 'comment': body5a+body5c+body5o}
    ]
    
    
    for job in job_list:
        if job['comment']:
            job['status'] = 'Error'
        else:
            job['status'] = 'Finished'
    
    
    
    job_list.append({'job': 'ZCO_4433_EPS_ICBAUTOBILL_UMIN2', 'status': co_job.sm37('ZCO_4433_EPS_ICBAUTOBILL_UMIN2'), 'comment': ''})
    job_list.append({'job': 'ZCO_4433_EPS_ICBAUTOBILL_UMIN1', 'status': co_job.sm37('ZCO_4433_EPS_ICBAUTOBILL_UMIN1'), 'comment': ''})
    job_list.append({'job': '/SIE/E_CO_E02525_PROCESS_CJEN', 'status': co_job.sm37('/SIE/E_CO_E02525_PROCESS_CJEN'), 'comment': ''})
    job_list.append({'job': 'SBT_CO_0010_AJ_PC1_COPA_INC_ORDE', 'status': co_job.sm37('SBT_CO_0010_AJ_PC1_COPA_INC_ORDE'), 'comment': ''})
    job_list.append({'job': 'SBT_CO_0010_AJ_PC1_W2W_UPDATE', 'status':co_job.sm37('SBT_CO_0010_AJ_PC1_COPA_INC_ORDE'), 'comment': ''})
    job_list.append({'job': 'SCO_CO0038_SII_ROLLUP_PCT_RECON', 'status': co_job.sm37('SCO_CO0038_SII_ROLLUP_PCT_RECON'), 'comment': ''})
    job_list.append({'job': 'SRE_PS_RA_01CX_4472', 'status': co_job.sm37('SRE_PS_RA_01CX_4472'), 'comment': ''})
    job_list.append({'job': 'SRE_PS_ST_01CX_4472', 'status': co_job.sm37('SRE_PS_ST_01CX_4472'), 'comment': ''})
    job_list.append({'job': 'ZHYP_DAILY_FINANCIALDATAEXTR', 'status': co_job.sm37('ZHYP_DAILY_FINANCIALDATAEXTR'), 'comment': ''})
    
    # Read the excel sheet into a pandas dataframe
    df = pd.read_excel(r'C:\\UserData\z004mfjs\Documents\sm37\co\spool.xlsx')



    ## Create a list to store the job names and statuses
    jobs = []

    # Iterate over the rows in the DataFrame and append the job names and statuses to the list
    for index, row in df.iterrows():
        if row["Status"] == "Canceled" or row["Status"] == "Active":
            jobs.append(f"{row['JobName']} - {row['Status']}")

    # Print the list of jobs
    #print(jobs)
    jobs = list(dict.fromkeys(jobs))
    jobs = "<br>".join(map(str, jobs))
    
    job_list.append({'job': 'RA & Settlements Logs Upload ( *SPOOL* )', 'status':'' ,'comment':jobs})
    
    job_list.append({'job': '5620 Lean ICB', 'status': 'Finished', 'comment': str(code_5620['Number of Entries'].sum())+lean })
    job_list.append({'job': '8620 Lean ICB', 'status': 'Finished', 'comment': str(code_8620['Number of Entries'].sum())+lean})
    job_list.append({'job': '9461 Lean ICB', 'status': 'Finished', 'comment':str(code_9461['Number of Entries'].sum())+lean})
    job_list.append({'job': '9462 Lean ICB', 'status': 'Finished', 'comment': str(code_9462['Number of Entries'].sum())+lean})

    
    html = '<table style="border-collapse: collapse; border: 1px solid black;">\n'
    html += '<tr>\n'
    html += '<tr style="background-color: yellow;">\n'
    html += f'<th style="border: 1px solid black; padding: 8px;" colspan="5">E1P.100 JOBS {today}</th>\n'
    html += '</tr>\n'
    html += '</thead>\n'
    
    html += '<tr style="background-color: yellow;">\n'
    html += '<th style="border: 1px solid black; padding: 8px;">S.No</th>\n'
    html += '<th style="border: 1px solid black; padding: 8px;">Name</th>\n'
    html += '<th style="border: 1px solid black; padding: 8px; background-color: yellow;">Status</th>\n'
    html += '<th style="border: 1px solid black; padding: 8px; background-color: yellow;">Status (Red/Yellow/Green)</th>\n'
    html += '<th style="border: 1px solid black; padding: 8px;">Comment</th>\n'
    html += '</tr>\n'
    html += '</thead>\n'
    html += '<tbody>\n'
    
    for i, job in enumerate(job_list, 1):
        if job['status'] == 'Finished':
            status_color = 'green'
        elif job['status'] == 'Error':
            status_color = 'yellow'
        else:
            status_color = 'red'
        html += '<tr>\n'
        html += f'<td style="border: 1px solid black; padding: 8px;">{i}</td>\n'
        html += f'<td style="border: 1px solid black; padding: 8px;">{job["job"]}</td>\n'
        html += f'<td style="border: 1px solid black; padding: 8px; color: black;">{job["status"]}</td>\n'
        html += f'<td style="border: 1px solid black; padding: 8px;"><b><span style="font-size:20.0pt;font-family:Wingdings;mso-fareast-font-family:Calibri;mso-fareast-theme-font:minor-latin;color:{status_color}">l<o:p></o:p></span></b></td>\n'
  
        html += f'<td style="border: 1px solid black; padding: 8px;">{job["comment"]}</td>\n'
        html += '</tr>\n'
    
    html += '</tbody>\n'
    html += '</table>\n'
    
    # create outlook email
    outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = 'mail@me.com'
    mail.CC = 'mail@me.com'
    mail.Subject = ' Job Monitoring CO ( E1P- 100 / E1P - 022 )'
    #mail.Attachments.Add(Source=r"C:\UserData\z004mfjs\Documents\abc.xlsx")
    
    # modify the mail body as per need
    mail.BodyFormat = 2
    
    intro = "<p>Hello Team,<br> As per provided revised job list. We have monitored jobs.<br>"  
    
    bye = "<br>  <br>  <br>Thanks and regards, <br> Srikanth <br>"
    break1="<br>"
    mail.HTMLBody = html
    
    #mail.Display()
    y=html


#part2= mail100()
#mail.HTMLBody = intro+bodyf+ bodya+bodyc+bodyr+body3+bye 
#mail.Send()
#remove this before the timing starts 



#main()
#co_job.read_jobs(0)
#combiner() 
#co_job.lean()
#co_job.spool('*spool*')

    


import win32com.client 
import pandas as pd
import time
import datetime
from dateutil.relativedelta import relativedelta
import calendar
from IPython.display import display, HTML
import e1p_022
def combiner022():
    import pandas as pd

    # Read the Excel file into a DataFrame
    df = pd.read_excel('022JOBS.xlsx', sheet_name='Sheet1', header=None)

    # Find the column index of the job name





    #s=read.sort_values('STATUS',ascending=(False))
    filename = (r"C:\UserData\z004mfjs\Documents\sm37\co\022CO_jobs_status.xlsx")
    excel=pd.read_excel(filename,index_col=0)
     #print(excel)
    list2 = excel.index.values.tolist()

    excel1=pd.read_excel(filename,index_col=1)
    list3= excel1.index.values.tolist()


    lsparent=[]
    lsjob=[]
    stat=[]

    for i in range(len(list2)):
        job_name= list2[i]
        status=list3[i]
        string_cols = df.select_dtypes(include='object').columns
        string_contains = df[string_cols].apply(lambda x: x.str.contains(job_name, na=False))
        col_index = string_contains.any().idxmax()

        # Get the header of the column index
        header = df.iloc[0, col_index] 
        print(job_name,header)
        lsparent.append(header)
        lsjob.append(job_name)
        stat.append(status)
        


    # print(df)
     #print(df.shape[0])
    df= df.dropna()
    df.to_clipboard(index=False, header=False)
    #print(df)
    df=pd.DataFrame()
    df['Parent']=lsparent
    df['Job name']=lsjob
    df['Status']=stat

    df.to_excel('022_combined.xlsx',index= False)


    
#combiner()


def mail022():
   
    df1 = pd.read_excel(r'C:\UserData\z004mfjs\Documents\sm37\co\022_combined.xlsx')
                    
    html_table = df1.to_html(index=False)
    #print(df1)
    par=df1["Parent"].values
    jobn=df1["Job name"].values
    status=df1["Status"].values
    
    data = {
        "parent":par,
      "JOB NAME": jobn,
      "STATUS": status
    }
    
    x = data["STATUS"]
    y= data["JOB NAME"]
    z=data["parent"]
    
    f=[]
    a=[]
    c=[]
    r=[]
    others=[]
    
    
    f_Z0010 =[]   
    a_Z0010=[]
    c_Z0010=[]
    others_Z0010=[]
    
    f_Z415S = []   
    a_Z415S = []
    c_Z415S = []
    others_Z415S = []
    
    f_Z543G = []   
    a_Z543G = []
    c_Z543G = []
    others_Z543G = []
    
    f_Z5520 = []   
    a_Z5520 = []
    c_Z5520 = []
    others_Z5520 = []
    
    f_Z5525 = []   
    a_Z5525 = []
    c_Z5525 = []
    others_Z5525 = []
    
    f_Z5530 = []   
    a_Z5530 = []
    c_Z5530 = []
    others_Z5530 = []
    
    f_Z5567 = []   
    a_Z5567 = []
    c_Z5567 = []
    others_Z5567 = []
    
    f_Z570M = []   
    a_Z570M = []
    c_Z570M = []
    others_Z570M = []

    

    for i in range(len(x)):
        if z[i] == "Z0010":
            if x[i] == "Finished":
                f_Z0010.append(y[i])
            elif x[i] == "Active":
                a_Z0010.append(y[i])
            elif x[i] == "Canceled":
                c_Z0010.append(y[i])
            else:
                others_Z0010.append(y[i])
        
        if z[i] == "Z415S":
            if x[i] == "Finished":
                f_Z415S.append(y[i])
            elif x[i] == "Active":
                a_Z415S.append(y[i])
            elif x[i] == "Canceled":
                c_Z415S.append(y[i])
            else:
                others_Z415S.append(y[i])
        
        elif z[i] == "Z543G":
            if x[i] == "Finished":
                f_Z543G.append(y[i])
            elif x[i] == "Active":
                a_Z543G.append(y[i])
            elif x[i] == "Canceled":
                c_Z543G.append(y[i])
            else:
                others_Z543G.append(y[i])
        elif z[i] == "Z5520":
            if x[i] == "Finished":
                f_Z5520.append(y[i])
            elif x[i] == "Active":
                a_Z5520.append(y[i])
            elif x[i] == "Canceled":
                c_Z5520.append(y[i])
            else:
                others_Z5520.append(y[i])
        
        elif z[i] == "Z5525":
            if x[i] == "Finished":
                f_Z5525.append(y[i])
            elif x[i] == "Active":
                a_Z5525.append(y[i])
            elif x[i] == "Canceled":
                c_Z5525.append(y[i])
            else:
                others_Z5525.append(y[i])
        
        elif z[i] == "Z5530":
            if x[i] == "Finished":
                f_Z5530.append(y[i])
            elif x[i] == "Active":
                a_Z5530.append(y[i])
            elif x[i] == "Canceled":
                c_Z5530.append(y[i])
            else:
                others_Z5530.append(y[i])
       
        elif z[i] == "Z5567":
            if x[i] == "Finished":
                f_Z5567.append(y[i])
            elif x[i] == "Active":
                a_Z5567.append(y[i])
            elif x[i] == "Canceled":
                c_Z5567.append(y[i])
            else:
                others_Z5567.append(y[i])
        
        elif z[i] == "Z570M":
            if x[i] == "Finished":
                f_Z570M.append(y[i])
            elif x[i] == "Active":
                a_Z570M.append(y[i])
            elif x[i] == "Canceled":
                c_Z570M.append(y[i])
            else:
                others_Z570M.append(y[i])
                
                
        # Z0010
    # removing duplicates
    f_Z0010 = list(dict.fromkeys(f_Z0010))
    a_Z0010 = list(dict.fromkeys(a_Z0010))
    c_Z0010 = list(dict.fromkeys(c_Z0010))
    others_Z0010 = list(dict.fromkeys(others_Z0010))
    
    # making the list to print one below the other
    pd.options.display.max_rows = 999999999
    testf_Z0010 = "<br>".join(map(str, f_Z0010))
    testa_Z0010 = "<br>".join(map(str, a_Z0010))
    testc_Z0010 = "<br>".join(map(str, c_Z0010))
    testo_Z0010 = "<br>".join(map(str, others_Z0010))
    
    # Z415S
    # removing duplicates
    f_Z415S = list(dict.fromkeys(f_Z415S))
    a_Z415S = list(dict.fromkeys(a_Z415S))
    c_Z415S = list(dict.fromkeys(c_Z415S))
    others_Z415S = list(dict.fromkeys(others_Z415S))
    
    # making the list to print one below the other
    testf_Z415S = "<br>".join(map(str, f_Z415S))
    testa_Z415S = "<br>".join(map(str, a_Z415S))
    testc_Z415S = "<br>".join(map(str, c_Z415S))
    testo_Z415S = "<br>".join(map(str, others_Z415S))
    
    # Z543G
    # removing duplicates
    f_Z543G = list(dict.fromkeys(f_Z543G))
    a_Z543G = list(dict.fromkeys(a_Z543G))
    c_Z543G = list(dict.fromkeys(c_Z543G))
    others_Z543G = list(dict.fromkeys(others_Z543G))

    # making the list to print one below the other
    testf_Z543G = "<br>".join(map(str, f_Z543G))
    testa_Z543G = "<br>".join(map(str, a_Z543G))
    testc_Z543G = "<br>".join(map(str, c_Z543G))
    testo_Z543G = "<br>".join(map(str, others_Z543G))
    
    # Z5520
    # removing duplicates
    f_Z5520 = list(dict.fromkeys(f_Z5520))
    a_Z5520 = list(dict.fromkeys(a_Z5520))
    c_Z5520 = list(dict.fromkeys(c_Z5520))
    others_Z5520 = list(dict.fromkeys(others_Z5520))
    
    # making the list to print one below the other
    testf_Z5520 = "<br>".join(map(str, f_Z5520))
    testa_Z5520 = "<br>".join(map(str, a_Z5520))
    testc_Z5520 = "<br>".join(map(str, c_Z5520))
    testo_Z5520 = "<br>".join(map(str, others_Z5520))
    
    # Z5525
    # removing duplicates
    f_Z5525 = list(dict.fromkeys(f_Z5525))
    a_Z5525 = list(dict.fromkeys(a_Z5525))
    c_Z5525 = list(dict.fromkeys(c_Z5525))
    others_Z5525 = list(dict.fromkeys(others_Z5525))
    
    # making the list to print one below the other
    testf_Z5525 = "<br>".join(map(str, f_Z5525))
    testa_Z5525 = "<br>".join(map(str, a_Z5525))
    testc_Z5525 = "<br>".join(map(str, c_Z5525))
    testo_Z5525 = "<br>".join(map(str,others_Z5525))
    
    
    # Z5530
    # removing duplicates
    f_Z5530 = list(dict.fromkeys(f_Z5530))
    a_Z5530 = list(dict.fromkeys(a_Z5530))
    c_Z5530 = list(dict.fromkeys(c_Z5530))
    others_Z5530 = list(dict.fromkeys(others_Z5530))
    
    # making the list to print one below the other
    pd.options.display.max_rows = 999999999
    testf_Z5530 = "<br>".join(map(str, f_Z5530))
    testa_Z5530 = "<br>".join(map(str, a_Z5530))
    testc_Z5530 = "<br>".join(map(str, c_Z5530))
    testo_Z5530 = "<br>".join(map(str, others_Z5530))
    
    # Z5567
    # removing duplicates
    f_Z5567 = list(dict.fromkeys(f_Z5567))
    a_Z5567 = list(dict.fromkeys(a_Z5567))
    c_Z5567 = list(dict.fromkeys(c_Z5567))
    others_Z5567 = list(dict.fromkeys(others_Z5567))
    
    # making the list to print one below the other
    testf_Z5567 = "<br>".join(map(str, f_Z5567))
    testa_Z5567 = "<br>".join(map(str, a_Z5567))
    testc_Z5567 = "<br>".join(map(str, c_Z5567))
    testo_Z5567 = "<br>".join(map(str, others_Z5567))
    
    # Z570M
    # removing duplicates
    f_Z570M = list(dict.fromkeys(f_Z570M))
    a_Z570M = list(dict.fromkeys(a_Z570M))
    c_Z570M = list(dict.fromkeys(c_Z570M))
    others_Z570M = list(dict.fromkeys(others_Z570M))
    
    # making the list to print one below the other
    testf_Z570M = "<br>".join(map(str, f_Z570M))
    testa_Z570M = "<br>".join(map(str, a_Z570M))
    testc_Z570M = "<br>".join(map(str, c_Z570M))
    testo_Z570M = "<br>".join(map(str, others_Z570M))
    
    import datetime

    today = datetime.date.today()
    first_day = datetime.date(today.year, today.month, 1)
    
    working_days = []
    for i in range(31):
        day = first_day + datetime.timedelta(days=i)
        if len(working_days) < 7 and day.weekday() < 5:
            working_days.append(day.day)
        elif len(working_days) == 7:
            break
    
    print("The first 7 working days of the month are:", working_days)
    
    #print(working_days[0])
    
    day=today.day
    
    print("today is : "+ str(day))

    if f_Z0010:
        bodyf = "<br>these are with status <b> Finished</b>:<br>" + " <span style='color: blue'>%s</span>" % testf_Z0010
    else:
        bodyf= ""
    if a_Z0010:
        bodya = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % testa_Z0010
    else:
        bodya = ""
    if c_Z0010:
        bodyc = "<br>these are with status <b> Canceled</b>:<br>" + " <span style='color: blue'>%s</span>" % testc_Z0010
    else:
        bodyc = ""
    if others_Z0010:
        if working_days[2] <= day <= working_days[5]:
            bodyo = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
            bodyo = "<br>these are with other status:<br>" + " <span style='color: blue'>%s</span>" % testo_Z0010
    else:
        bodyo = ""
 
        
    if f_Z415S:
        body1f = "<br>these are with status <b> Finished</b>:<br>" + " <span style='color: blue'>%s</span>" % testf_Z415S
    else:
        body1f= ""
    if a_Z415S:
        bodya1 = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % testa_Z415S
    else:
        body1a = ""
    if c_Z415S:
        body1c = "<br>these are with status <b> Canceled</b>:<br>" + " <span style='color: blue'>%s</span>" % testc_Z415S
    else:
        body1c = ""
    if others_Z415S:
        if working_days[2] <= day <= working_days[5]:
            bod1yo = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
            body1o = "<br>these are with other status:<br>" + " <span style='color: blue'>%s</span>" % testo_Z415S
    else:
        body1o = ""
        
        
    if f_Z543G:
        body1f_Z543G = "<br>these are with status <b> Finished</b>:<br>" + " <span style='color: blue'>%s</span>" % testf_Z543G
    else:
        body1f_Z543G= ""
    if a_Z543G:
        body1a_Z543G = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % testa_Z543G
    else:
        body1a_Z543G = ""
    if c_Z543G:
        body1c_Z543G = "<br>these are with status <b> Canceled</b>:<br>" + " <span style='color: blue'>%s</span>" % testc_Z543G
    else:
        body1c_Z543G = ""
    if others_Z543G:
        if working_days[2] <= day <= working_days[5]:
            body1o_Z543G = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
            body1o_Z543G = "<br>these are with other status:<br>" + " <span style='color: blue'>%s</span>" % testo_Z543G
    else:
        body1o_Z543G = ""
        
        
    if f_Z5520:
        body1f_Z5520 = "<br>these are with status <b> Finished</b>:<br>" + " <span style='color: blue'>%s</span>" % testf_Z5520
    else:
        body1f_Z5520= ""
    if a_Z5520:
        body1a_Z5520 = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % testa_Z5520
    else:
        body1a_Z5520 = ""
    if c_Z5520:
        body1c_Z5520 = "<br>these are with status <b> Canceled</b>:<br>" + " <span style='color: blue'>%s</span>" % testc_Z5520
    else:
        body1c_Z5520 = ""
    if others_Z5520:
        if working_days[2] <= day <= working_days[5]:
            body1o_Z5520 = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
            body1o_Z5520 = "<br>these are with other status:<br>" + " <span style='color: blue'>%s</span>" % testo_Z5520
    else:
        body1o_Z5520 = ""
        
        
    if f_Z5525:
        body1f_Z5525 = "<br>these are with status <b> Finished</b>:<br>" + " <span style='color: blue'>%s</span>" % testf_Z5525
    else:
        body1f_Z5525= ""
    if a_Z5525:
        body1a_Z5525 = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % testa_Z5525
    else:
        body1a_Z5525 = ""
    if c_Z5525:
        body1c_Z5525 = "<br>these are with status <b> Canceled</b>:<br>" + " <span style='color: blue'>%s</span>" % testc_Z5525
    else:
        body1c_Z5525 = ""
    if others_Z5525:
        if working_days[2] <= day <= working_days[5]:
            body1o_Z5525 = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
            body1o_Z5525 = "<br>these are with other status:<br>" + " <span style='color: blue'>%s</span>" % testo_Z5525
    else:
        
        body1o_Z5525 = ""
        
        
    if f_Z5530:
        body2f = "<br>these are with status <b> Finished</b>:<br>" + " <span style='color: blue'>%s</span>" % testf_Z5530
    else:
        body2f= ""
    if a_Z5530:
        body2a = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % testa_Z5530
    else:
        body2a = ""
    if c_Z5530:
        body2c = "<br>these are with status <b> Canceled</b>:<br>" + " <span style='color: blue'>%s</span>" % testc_Z5530
    else:
        body2c = ""
    if others_Z5530:
        if working_days[2] <= day <= working_days[5]:
            body2o = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
            body2o = "<br>these are with other status:<br>" + " <span style='color: blue'>%s</span>" % testo_Z5530
    else:
        body2o = ""
    
    if f_Z5567:
        body21f = "<br>these are with status <b> Finished</b>:<br>" + " <span style='color: blue'>%s</span>" % testf_Z5567
    else:
        body21f= ""
    if a_Z5567:
        body21a = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % testa_Z5567
    else:
        body21a = ""
    if c_Z5567:
        body21c = "<br>these are with status <b> Canceled</b>:<br>" + " <span style='color: blue'>%s</span>" % testc_Z5567
    else:
        body21c = ""
    if others_Z5567:
        if working_days[2] <= day <= working_days[5]:
            body21o = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
            body21o = "<br>these are with other status:<br>" + " <span style='color: blue'>%s</span>" % testo_Z5567
    else:
        body21o = ""
    
    if f_Z570M:
        body3f = "<br>these are with status <b> Finished</b>:<br>" + " <span style='color: blue'>%s</span>" % testf_Z570M
    else:
        body3f= ""
    if a_Z570M:
        body3a = "<br>these are with status <b> Active</b>:<br>" + " <span style='color: blue'>%s</span>" % testa_Z570M
    else:
        body3a = ""
    if c_Z570M:
        body3c = "<br>these are with status <b> Canceled</b>:<br>" + " <span style='color: blue'>%s</span>" % testc_Z570M
    else:
        body3c = ""
    if others_Z570M:
        if working_days[2] <= day <= working_days[5]:
            body3o = '<span style="background:yellow;mso-highlight:yellow">Job will successfully run after U + 7</span>'
        else:
            body3o = "<br>these are with other status:<br>" + " <span style='color: blue'>%s</span>" % testo_Z570M
    else:
        body3o = ""

    

    
    job_list = [
        {'job': 'AE/0042 Job ZCO_0010_PAPCA_COMP', 'status': e1p_022.sm37('ZCO_0010_PAPCA_COMP'), 'comment': ''},
        {'job': 'AE/0042 Job ZCO_0010_PRJ_ALLWBS_SET', 'status': e1p_022.sm37('ZCO_0010_PRJ_ALLWBS_SET'), 'comment':'' },
        {'job': 'AE/0042 Job ZCO_0010_PRJ_ALLWBS_IOC', 'status': e1p_022.sm37('ZCO_0010_PRJ_ALLWBS_IOC'), 'comment': ''},
        {'job': 'AE/0042 Job ZCO_0010_SDI_ALLPLANTS_SET', 'status': e1p_022.sm37('ZCO_0010_SDI_ALLPLANTS_SET'), 'comment': ''},
        {'job': 'AE/0042 Job ZCO_0010_PRD_ALLPLANTS_VAR', 'status': e1p_022.sm37('ZCO_0010_PRD_ALLPLANTS_VAR'), 'comment': ''},
        {'job': 'AE/0042 Job ZCO_0010_PRD_ALLPLANTS_OHC', 'status': e1p_022.sm37('ZCO_0010_PRD_ALLPLANTS_OHC'), 'comment': ''},
        {'job': 'AE/0042 Job ZCO_0010_PRD_ALLPLANTS_WIP', 'status': e1p_022.sm37('ZCO_0010_PRD_ALLPLANTS_WIP'), 'comment': ''},
        
        
        {'job': '0010 RA/Settlement/Variance Jobs', 'status': '', 'comment': bodya+bodyc+bodyo},
        {'job': '415S RA/Variance/Settlement/OHC', 'status': '', 'comment': body1a+body1c+body1o},
        {'job': '543G RA/Settlement/Variance/OHC', 'status': '', 'comment':  body1a_Z543G+ body1c_Z543G+ body1o_Z543G},
        {'job': '5520 RA/Settlement/Variance/OHC', 'status': '', 'comment':  body1a_Z5520+ body1c_Z5520+ body1o_Z5520},
        {'job': '5525 RA/Settlement/Variance/OHC', 'status': '', 'comment':  body1a_Z5525+body1c_Z5525+body1o_Z5525},
        {'job': '5530 RA/Settlement/Variance/OHC', 'status': '', 'comment': body2a+body2c+body2o},
        {'job': '5567 RA/Settlement/Variance/OHC', 'status': '', 'comment': body21a+body21c+body21o},
        {'job': '570M RA/Settlement/Variance/OHC', 'status': '', 'comment': body3a+body3c+body3o}
    ]
    
    #for job in job_list:
        #print(f"{job['job']} - {job['status']}")
    '''
    for job in job_list:
        if job['comment']:
            job['status'] = 'Error'
        else:
            job['status'] = 'Finished'
    '''
    
    
    
    
    
    html = '<table style="border-collapse: collapse; border: 1px solid black;">\n'
    html += '<tr>\n'
    html += '<thead style="background-color: yellow;">\n'
    html += f'<th style="border: 1px solid black; padding: 8px;" colspan="5">E1P.022 JOBS {today}</th>\n'
    html += '</tr>\n'
    html += '</thead>\n'
    
    html += '<tr style="background-color: yellow;">\n'
    html += '<th style="border: 1px solid black; padding: 8px;">S.No</th>\n'
    html += '<th style="border: 1px solid black; padding: 8px;">Name</th>\n'
    html += '<th style="border: 1px solid black; padding: 8px; background-color: yellow;">Status</th>\n'
    html += '<th style="border: 1px solid black; padding: 8px; background-color: yellow;">Status (Red/Yellow/Green)</th>\n'
    html += '<th style="border: 1px solid black; padding: 8px;">Comment</th>\n'
    html += '</tr>\n'
    html += '</thead>\n'
    html += '<tbody>\n'
    
    for i, job in enumerate(job_list, 1):
        if job['status'] == 'Finished':
            status_color = 'green'
        elif job['status'] == 'Error':
            status_color = 'yellow'
        else:
            status_color = 'red'
        html += '<tr>\n'
        html += f'<td style="border: 1px solid black; padding: 8px;">{i}</td>\n'
        html += f'<td style="border: 1px solid black; padding: 8px;">{job["job"]}</td>\n'
        html += f'<td style="border: 1px solid black; padding: 8px; color: black;">{job["status"]}</td>\n'
        
        html += f'<td style="border: 1px solid black; padding: 8px;"><b><span style="font-size:20.0pt;font-family:Wingdings;mso-fareast-font-family:Calibri;mso-fareast-theme-font:minor-latin;color:{status_color}">l<o:p></o:p></span></b></td>\n'
  
        html += f'<td style="border: 1px solid black; padding: 8px;">{job["comment"]}</td>\n'
        html += '</tr>\n'
    
    html += '</tbody>\n'
    html += '</table>\n'
    
    # create outlook email
    outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = 'mail@me.com'
    mail.CC = 'mail@me.com'
    mail.Subject = ' Job Monitoring CO ( E1P- 100 / E1P - 022 )'
    #mail.Attachments.Add(Source=r"C:\UserData\z004mfjs\Documents\abc.xlsx")
    
    # modify the mail body as per need
    mail.BodyFormat = 2
    
    intro = "<p>Hello Team,<br> As per provided revised job list. We have monitored jobs.<br>"  
    
    bye = "<br>  <br>  <br>Thanks and regards, <br> Srikanth <br>"
    break1="<br>"
    mail.HTMLBody = html
    
    #ail.Display()
    
    z=html
    
#combiner022()    
#mail022()
 
#part3 = mail022()




def mail():       
    outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = 'mail@me.com'
    mail.CC = 'mail@me.com'
    mail.Subject = ' Job Monitoring CO ( E1P- 100 / E1P - 022 )'
    #mail.Attachments.Add(Source=r"C:\UserData\z004mfjs\Documents\abc.xlsx")
    
    # modify the mail body as per need
    mail.BodyFormat = 2
    
    intro = "<p>Hello Team,<br> As per provided revised job list. We have monitored jobs.<br>"  
    
    bye = "<br>  <br>  <br>Thanks and regards, <br> Srikanth <br>"
    break1="<br>"
    mail.HTMLBody = temp()#+"<br>""<br>""<br>"+y+"<br>""<br>""<br>"+ z
    mail.Display()


if __name__ == "__main__":
    co_job.login()

    #co_job.read_jobs(0)
    #combiner() 
    #co_job.lean()
    #co_job.spool('*spool*')
    temp() 
    mail100()
    co_job.close() 
    
    e1p_022.login()
    #e1p_022.read_jobs(0)   
    #combiner022()   
    mail022()
    mail()
    
        

    
