import json
import numpy
import xlsxwriter
import glob
import getpass
from urlparse import urlparse
import os
from datetime import datetime
import optparse
from selenium import webdriver
import time

defaultLoc = '/Users/%s/Desktop/' % getpass.getuser()

os.system('clear')
parser = optparse.OptionParser()
parser.add_option('-i', '--iterations', help = 'Number of times to iterate, i.e: loop 5 times', dest = 'iterations', default = 1, type = 'int')
parser.add_option('-r', '--repeat', help = 'Repeat frequency in seconds, i.e: run every 60 seconds', dest = 'frequency', default = 1, type = 'int')
parser.add_option('-u', '--url', help = 'Web page to create HAR and Charts for', dest = 'url', default = None, type = 'string')
parser.add_option('-n', '--username', help = 'User name to use for site login', dest = 'userName', default = None, type = 'string')
parser.add_option('-o', '--output_location', help = 'Location to write Excel output file, if no location given than write to the user\'s desktop', dest = 'excelOutput', default = defaultLoc, type = 'string')
parser.add_option('-v', '--verbose', help = 'Verbose option prints detailed request information to the screen', dest = 'verbose', default = False, action = 'store_true')
(opts, args) = parser.parse_args()

if opts.url is None:
    print 'Error: *** No URL provided ***. Please supply the url using the -u or --url option'
    exit(-1)
elif opts.userName is None:
    print 'Error: *** No username was provided for application login ***. Please supply application username using the -n or --username option'
    exit(-1)

print 'Starting HAR to Perato Generator with %d iterations executing every %d seconds...' % (opts.iterations, opts.frequency)

fireBugPath = '/Users/Greg/HarViewer/firebug-2.0.9-fx.xpi'
netExportPath = '/Users/Greg/HarViewer/netExport-0.8.xpi'

theUrl = urlparse(opts.url)
password = 'Password99'

 # Create profile and add Firefox extensions
profile = webdriver.FirefoxProfile()
profile.add_extension(fireBugPath)
profile.add_extension(netExportPath)
profile.set_preference("app.update.enabled", "false")

# Set FireBug Preferences
profile.set_preference("extensions.firebug.currentVersion", "2.0.9")
profile.set_preference("extensions.firebug.net.enableSites", "true")
profile.set_preference("extensions.firebug.allPagesActivation", "on")
profile.set_preference("extensions.firebug.defaultPanelName", "net")

# Set NetExport Preferences 
profile.set_preference("extensions.firebug.netexport.alwaysEnableAutoExport", "true")
profile.set_preference("extensions.firebug.netexport.showPreview", "false")
profile.set_preference("extensions.firebug.netexport.defaultLogDir", "/Users/Greg/HarViewer/hars")

ff = webdriver.Firefox(firefox_profile=profile)
ff.get('http://' + theUrl.netloc)

login = ff.find_element_by_id('login')
passWord = ff.find_element_by_id('password')
btnSignin = ff.find_element_by_css_selector('a.btn')

login.send_keys('greg.buckner')
passWord.send_keys(password)
btnSignin.click()

time.sleep(10)

oldHARS = glob.glob('/Users/Greg/HarViewer/hars/*.har')

for deleteHAR in range(0,len(oldHARS)):
    os.remove(oldHARS[deleteHAR]) 
    print 'Cleaning up old HAR file... %s' % oldHARS[deleteHAR]

timeStamp = datetime.now()
appendTime = timeStamp.strftime('%m-%d-%y+%H:%M:%S')
workBookName = 'Performance_Perato+%s.xlsx' % (appendTime)
workbook = xlsxwriter.Workbook(opts.excelOutput + '/' + workBookName)

for iteration in range(1, opts.iterations + 1):
    iterationMsg = 'Processing iteration %d of %d...' % (iteration, opts.iterations)
    print'\n'
    print '*' * len(iterationMsg)
    print iterationMsg 
    print '*' * len(iterationMsg)

    #ff.get('about:blank')
    ff.get(opts.url)
    time.sleep(10)


    harFileName = glob.glob('/Users/Greg/HarViewer/hars/*.har')

    with open(harFileName[0], 'r') as jsonFile:
        harObj = json.load(jsonFile)
    

    os.remove(harFileName[0])
    #print 'Deleted current HAR file... %s' % harFileName[0]

    totTime = []
    culTotTime = []
    urlReqs = []
    requestNums = []
    cellComments = []

    for i in range(0, len(harObj['log']['entries'])):
        requestNums.append('Request %d' % (i + 1))
        urlReqs.append(harObj['log']['entries'][i]['request']['url'])
        total =+ harObj['log']['entries'][i]['timings']['receive'] + \
        harObj['log']['entries'][i]['timings']['send'] + \
        harObj['log']['entries'][i]['timings']['connect'] + \
        harObj['log']['entries'][i]['timings']['dns'] + \
        harObj['log']['entries'][i]['timings']['blocked'] + \
        harObj['log']['entries'][i]['timings']['wait']
        #print '%s %s' % (max(harObj['log']['entries'][i]['timings'], key = harObj['log']['entries'][i]['timings'].get), max(harObj['log']['entries'][i]['timings'].values()))
        maxTime = max(harObj['log']['entries'][i]['timings'].values())
        pctTime = float((float(maxTime) / float(total)) * 100)
        if opts.verbose == True:
            print 'Request %d spent %dms in %s state which took %d percent of the total response time receiving %d bytes' % \
            ((i + 1), maxTime, max(harObj['log']['entries'][i]['timings'], key = harObj['log']['entries'][i]['timings'].get).upper(), \
            pctTime, harObj['log']['entries'][i]['response']['bodySize'])
        totTime.append(total)
    
    print 'Longest response was request number %d - %s which took %.2f seconds total time' % ((totTime.index(max(totTime))+1), urlReqs[totTime.index(max(totTime))], max(totTime) / 1000.0)
    print 'Page onLoad was %.2f seconds' % (harObj['log']['pages'][0]['pageTimings']['onLoad'] / 1000.0)
    reqTimes = zip(requestNums, totTime)
    urlTimes = zip(urlReqs, totTime)
    #Sort by longest response time desc
    sortedByLongestTime = sorted(reqTimes, key=lambda tup: tup[1], reverse=True)
    sortedForComments = sorted(urlTimes, key=lambda tup: tup[1], reverse=True)

    #Prepare cell comments
    for cellCom in range(0, len(sortedForComments)):
                     cellComments.append(list(sortedForComments[cellCom])[0])

    workTable_Col1 = []
    workTable_Col2 = []
    workTable_Col3 = []
    workTable_Col4 = []

    
    for tup in range(0, len(sortedByLongestTime)):
                     workTable_Col1.append(list(sortedByLongestTime[tup])[0])
                     workTable_Col2.append(list(sortedByLongestTime[tup])[1])
                     workTable_Col4.append('0.8')

    workTable_Col3 = numpy.cumsum(workTable_Col2)
    
    #Calculate percentage
    percentage = []
    for p in range(0, len(totTime)):
        percentage.append(format(workTable_Col3[p] / float(workTable_Col3[-1]),'.2f'))

    percentage = map(float, percentage)
    workTable_Col4 = map(float, workTable_Col4)

    thisWorkSheet = 'Iteration%d' % iteration
    worksheet = workbook.add_worksheet(thisWorkSheet)

    #Create formats
    pctFormat = workbook.add_format({'num_format' : '0.0%'})
    bold = workbook.add_format({'bold' : True})

    worksheet.set_column('A:F', 16)

    columnChart = workbook.add_chart({'type' : 'column'})
    columnChart.set_chartarea({'border' : {'none' : True }})

    #Write to columns
    worksheet.write_row('A1', ['HTTP Requests', 'Response Time(ms)', 'Cumulative Time', 'Cumulative Pct', '80% Line'], bold)
    worksheet.write_column('A2', workTable_Col1)
    worksheet.write_column('B2', workTable_Col2)
    worksheet.write_column('C2', workTable_Col3)
    worksheet.write_column('D2', percentage, pctFormat)
    worksheet.write_column('E2', workTable_Col4, pctFormat)

    for cc in range(0, len(cellComments)):
        inputCell = 'A' + str(cc + 2)
        worksheet.write_comment(inputCell, cellComments[cc], {'height' : 80, 'width' : 240})

    #Chart series
    colChartSeries = '%s!$B$2:$B$%d' % (thisWorkSheet, len(workTable_Col2))
    lineChartSeries1 = '%s!$D$2:$D$%d' % (thisWorkSheet, len(workTable_Col4))
    lineChartSeries2 = '%s!$E$2:$E$%d' % (thisWorkSheet, len(percentage))
    chartCategories =  '%s!$A$2:$A$%d' % (thisWorkSheet, len(workTable_Col1))


    columnChart.add_series({'values' : colChartSeries, 'categories' : chartCategories})

    columnChart.set_legend({'position': 'none'})            
    chartTitle = urlparse(ff.current_url)
    chartTitle = chartTitle.path + ' | ' + 'Page onLoad %.2f seconds' % (harObj['log']['pages'][0]['pageTimings']['onLoad'] / 1000.0)
    columnChart.set_title({'name' : chartTitle.title()})

    columnChart.set_y_axis({'name': 'Request (ms)', 'min': 0, 'max': max(workTable_Col2), 'major_gridlines' : {'visible'   : False}})
    columnChart.set_y2_axis({'max' : 1,  'number' : 'percentage', 'num_format' : 'percentage'})

    lineChart = workbook.add_chart({'type' : 'line'})
    lineChart.add_series({'values' : lineChartSeries1, 'categories' : chartCategories,  'marker' : {'type' : 'automatic', 'size' : 8}, 'y2_axis' : 1, 'line' : {'color' : 'red'}})
    lineChart.add_series({'values' : lineChartSeries2, 'categories' : chartCategories, 'y2_axis' : 1, 'line' : {'color' : 'gray'}})

    columnChart.combine(lineChart)

    columnChart.set_size({'height' : 600, 'width' : 900})

    worksheet.insert_chart('E1', columnChart)
    
    if iteration == opts.iterations:
        print '\n'
        print 'All iterations finished...'
        workbook.close()
        print '\n'
        print 'Check Excel file %s/%s for results' % (opts.excelOutput, workBookName)
    else:
        print '\n'
        print 'Sleeping for %d seconds until next iteration...' % opts.frequency
        time.sleep(opts.frequency)
ff.quit()









