import datetime as dt
import requests

report = {
    "date" : dt.datetime.now().strftime('%Y%m%d'),
    "time" : dt.datetime.now().strftime('%H:%M:%S'),
    'taiway_yyyMMdd' : f"{dt.datetime.now().year - 1911}/{dt.datetime.now():%m}/{dt.datetime.now():%d}",
    "url" : "http://172.19.1.21/medpt/medptlogin.php",
    "url_connected" : False,
    "branch" : [{'vhwc':'wc'},{'vhcy':'cy'}],
    "item" : "Prescription_log",
    "png_foldername" : "d:\\mis\\",
    "png_filename" : None,
    'png_filepath' : None,
    'crop_images' : None,
    'fail_list' : None,
    "message" : None
}


for branch in report['branch']:
    for fullname, name in branch.items():
        print(fullname)
        print(name)
        
        
url = "http://172.19.1.21/medpt/cyp2001.php"
data = {'g_yyymmdd_s': 'taiwan_yyymmdd','from': 'b',}

response = requests.post(url=url,data=data)        
print(response.status_code)