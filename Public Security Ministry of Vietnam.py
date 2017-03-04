import datetime, re, sys, requests, xlsxwriter
from bs4 import BeautifulSoup

domain = 'http://www.mps.gov.vn'
baseURL = domain + '/c/portal/render_portlet'
payload = {
	'p_l_id':13515,
	'p_p_id':'vcmsviewcontent_INSTANCE_GbkG',
	'p_p_lifecycle':0,
	'p_p_state':'normal',
	'p_p_mode':'view',
	'p_p_col_id':'column-2',
	'p_p_col_pos':0,
	'p_p_col_count':1,
	'currentURL':'/web/guest/ct_trangchu/-/vcmsviewcontent/GbkG/2015/2015',
	'_vcmsviewcontent_INSTANCE_GbkG_categoryId':2015,
	'_vcmsviewcontent_INSTANCE_GbkG_cat_parent':2015,
	'_vcmsviewcontent_INSTANCE_GbkG_struts_action':'/vcmsviewcontent/view',
	'_vcmsviewcontent_INSTANCE_GbkG_page':'', # page id
	'_vcmsviewcontent_INSTANCE_GbkG_articleId':'', # artical id for each record
}
fileName = 'Public Security Ministry of Vietnam_' + datetime.datetime.now().strftime('%Y%m%d%H%M%S')
records = []

def processListPage(response):
	soup = BeautifulSoup(response.content, 'html.parser')
	items = soup.select_one('div.smallpage').findAll('div', {'class': 'clearfix'})
	for item in items:
		title = item.select_one('a.fon6')
		detail = item.select_one('div.fon5')
		record = {
			'title': 'Unknown' if title is None else title.find('b').text.strip(),
			'titleurl': '' if title is None else title['href'],
			'date': '' if title is None else title.findNext('span').text,
			'detail': re.sub('<.*?>', '', re.sub('<br.*?>', '\n', str(detail))).strip() 
		}
		global records
		records.append(record)
		
def getPageCount(response):
	soup = BeautifulSoup(response.content, 'html.parser')
	link = soup.findAll('a', {'class': 'next_article'})
	if link is None or len(link)<1 or not link[1].has_attr('onclick'):
		return 0
	else: 
		number = link[1]['onclick'].replace("_vcmsviewcontent_INSTANCE_GbkG_submitForm('", '').replace("','');", '')
		return int(number)

def downloadProgress(ratio):
	sys.stdout.write('\rDownload Data: ' + str(round(ratio*100)) + ' %')
	sys.stdout.flush()	

def saveData():
	wk = xlsxwriter.Workbook(fileName + '.xlsx')
	header_format = wk.add_format({'bold': True}) 
	cell_format = wk.add_format({'align': 'left','valign': 'vcenter'})

	ws = wk.add_worksheet()
	ws.write(0, 0, 'Title', header_format)
	ws.write(0, 1, 'Detail', header_format)
	row = 1
	for record in records:
		ws.write(row, 0, record['title'], cell_format)
		ws.write(row, 1, record['detail'], cell_format)
		row += 1
	wk.close()
	
def main():
	response = requests.post(url=baseURL, data=payload)
	pageCount = getPageCount(response)
	
	processListPage(response)
	downloadProgress(1/pageCount)

	for i in range(2, pageCount+1, 1):
		payload['_vcmsviewcontent_INSTANCE_GbkG_page'] = i
		response = requests.post(url=baseURL, data=payload)
		processListPage(response)
		downloadProgress(i/pageCount)

	print('\n Saving Data...')
	saveData()
	print('Script finished...')
	
if __name__ == '__main__':
	main()